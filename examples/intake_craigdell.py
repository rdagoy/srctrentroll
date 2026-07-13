"""Intake for Craigdell Gardens (AppFolio 'Rent Roll' export).

Sheet 'Sheet1'.  Header row 10; a property-group banner (row 11); units are the
rows whose Unit cell starts with "Unit "; grand totals near the bottom
("97 Units" / "Total 97 Units") are used to reconcile.

Columns: A Unit | B Tags | C BD/BA | D Tenant | E Status | F Sqft
  | G Market Rent | H Rent | I Deposit | J Move-in | K Lease From | L Lease To
  | M Move-out | N Monthly Charges

Mapping:
  * market_rent = G ; contract_rent = H (in-place; 0 if vacant)
  * other_income = N (Monthly Charges), shown as one OneLineRR column
  * BD/BA "b/ba" -> b BD / ba BA ; lease_start = K (Lease From)
  * Status: Current/Notice -> occupied; Vacant -> vacant. "DOWN" tag noted.
"""
import os
import re
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from rent_roll_processor import Unit, RollConfig, build_exhibits  # noqa: E402
from rent_roll_processor.naming import output_filename            # noqa: E402

PROPERTY = "Craigdell Gardens"
AS_OF = date(2026, 7, 13)
BDBA_RE = re.compile(r"^\s*(\d+)\s*/\s*([\d.]+)")


def decode(bdba):
    m = BDBA_RE.match(str(bdba or ""))
    if not m:
        return str(bdba or "").strip(), None, None
    beds = int(m.group(1))
    baths = float(m.group(2))
    fp = f"{beds} BD / {baths:g} BA"
    return f"{beds}/{baths:g}", fp, beds, baths


def num(v):
    return v if isinstance(v, (int, float)) else None


def load(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True)["Sheet1"]
    units = []
    for r in range(11, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not (isinstance(a, str) and a.strip().lower().startswith("unit ")):
            continue
        code, fp, beds, baths = decode(ws.cell(r, 3).value)
        status = str(ws.cell(r, 5).value or "").strip()
        tag = str(ws.cell(r, 2).value or "").strip()
        vacant = "vacant" in status.lower()   # Current/Notice -> occupied
        name = str(ws.cell(r, 4).value or "").strip()
        other = num(ws.cell(r, 14).value) or 0
        units.append(Unit.from_dict({
            "unit_id": a.strip()[5:].strip() or a.strip(),   # drop "Unit " prefix
            "unit_type": code,
            "floor_plan": fp,
            "beds": beds,
            "baths": baths,
            "sqft": num(ws.cell(r, 6).value),
            "tenant_name": name or ("Vacant" if vacant else None),
            "occupancy": "Vacant" if vacant else "Occupied",
            "market_rent": num(ws.cell(r, 7).value) or 0,             # G
            "contract_rent": 0 if vacant else (num(ws.cell(r, 8).value) or 0),  # H
            "other_income": other,
            "other_income_items": {"Monthly Charges": other} if other else {},
            "move_in": ws.cell(r, 10).value,
            "lease_start": ws.cell(r, 11).value,   # Lease From
            "lease_end": ws.cell(r, 12).value,     # Lease To
            "move_out": ws.cell(r, 13).value,
            "designation": tag or None,
        }))
    return units


def main():
    args = sys.argv[1:]
    src = args[0] if args else None
    out = args[1] if len(args) > 1 else output_filename(PROPERTY)
    if not src:
        print("usage: intake_craigdell.py <source.xlsx> [out.xlsx]")
        return 2
    units = load(src)
    config = RollConfig(property_name=PROPERTY, as_of_date=AS_OF, uw_market_rents={})
    build_exhibits(units, config, out)

    from rent_roll_processor.aggregate import is_occupied
    occ = sum(1 for u in units if is_occupied(u, config.model_occupied))
    vac = sum(1 for u in units if u.occupancy == "Vac")
    print(f"Wrote {out}")
    print(f"  units={len(units)}  occupied={occ}  vacant={vac}  occ%={occ/len(units):.2%}")
    print("  reconcile (mine vs source totals row):")
    for label, mine, src_tot in [
        ("Units", len(units), 97),
        ("Sq Ft", sum(u.sqft or 0 for u in units), 76200),
        ("Market Rent", sum(u.market_rent for u in units), 85460),
        ("Rent (in-place)", sum(u.contract_rent for u in units), 79227),
        ("Monthly Charges", sum(u.other_income for u in units), 1590),
    ]:
        flag = "OK" if abs(mine - src_tot) < 1 else "DIFF"
        print(f"    {label:16s} mine={mine:>10,.0f}  source={src_tot:>10,.0f}  [{flag}]")
    return 0


if __name__ == "__main__":
    sys.exit(main())
