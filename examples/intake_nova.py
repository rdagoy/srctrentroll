"""Intake for Nova (AppFolio 'Rent Roll', single property, academic-year lease-up).

Sheet 'Sheet1'.  A single property ("Property Groups: NOVA") -- not a portfolio,
so no per-property scoping.  Flat unit list under an 8-column header, closed by
a grand-total row (SF / Rent) used to reconcile.  Student-style by-the-bed
leasing shows up as split unit ids ("Unit 205 A" / "205 B"); each is its own row.

Columns: A Unit | B Unit Type | C Tenant | D Status | E Sqft | F Rent
  | G Lease From | H Lease To

  * unit type "<n>bed/<m>bath" -> n BD / m BA ; "Studio Apartment" -> 0 BD / 1 BA
  * Status: Leased / Pending -> occupied; Vacant-* -> vacant.  Unit 218 is
    "Tenant Pending" with committed rent, so it counts occupied.
  * No market-rent column -> market rent = in-place Rent (placeholder); vacant
    market is then derived from same-unit-type comps (#10).
"""
import os
import re
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from rent_roll_processor import Unit, RollConfig, build_exhibits  # noqa: E402
from rent_roll_processor.naming import output_filename            # noqa: E402

PROPERTY = "Nova"
AS_OF = date(2026, 9, 1)           # source header "As of: 09/01/2026"
BDBA_RE = re.compile(r"(\d+)\s*bed\s*/\s*(\d+)\s*bath", re.I)


def decode(utype):
    raw = str(utype or "").strip()
    if "studio" in raw.lower():
        return "0 BD / 1 BA", 0, 1
    m = BDBA_RE.search(raw)
    if m:
        b, ba = int(m.group(1)), int(m.group(2))
        return f"{b} BD / {ba} BA", b, ba
    return raw, None, None


def num(v):
    return v if isinstance(v, (int, float)) else None


def load(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True)["Sheet1"]
    units = []
    for r in range(9, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not (isinstance(a, str) and a.strip().lower().startswith("unit ")):
            continue                                  # skips header / total / blanks
        b = ws.cell(r, 2).value
        fp, beds, baths = decode(b)
        status = str(ws.cell(r, 4).value or "").strip()
        vacant = status.lower().startswith("vacant")  # Leased / Pending -> occupied
        name = str(ws.cell(r, 3).value or "").strip()
        rent = num(ws.cell(r, 6).value) or 0
        units.append(Unit.from_dict({
            "unit_id": a.strip()[5:].strip(),          # drop "Unit " prefix
            "unit_type": str(b).strip(),               # strict vacant-market comp key
            "floor_plan": fp,
            "beds": beds,
            "baths": baths,
            "sqft": num(ws.cell(r, 5).value),
            "tenant_name": name or ("Vacant" if vacant else None),
            "occupancy": "Vacant" if vacant else "Occupied",
            "market_rent": 0 if vacant else rent,      # no market col -> in-place placeholder
            "contract_rent": 0 if vacant else rent,
            "lease_start": ws.cell(r, 7).value,        # Lease From
            "lease_end": ws.cell(r, 8).value,          # Lease To
        }))
    return units


def main():
    args = sys.argv[1:]
    src = args[0] if args else None
    out = args[1] if len(args) > 1 else output_filename(PROPERTY)
    if not src:
        print("usage: intake_nova.py <source.xlsx> [out.xlsx]")
        return 2
    units = load(src)
    config = RollConfig(property_name=PROPERTY, as_of_date=AS_OF, uw_market_rents={})
    build_exhibits(units, config, out)

    from rent_roll_processor.aggregate import is_occupied
    occ = sum(1 for u in units if is_occupied(u, config.model_occupied))
    vac = sum(1 for u in units if u.occupancy == "Vac")
    print(f"Wrote {out}")
    print(f"  units={len(units)}  occupied={occ}  vacant={vac}  occ%={occ/len(units):.2%}")
    print("  reconcile (mine vs source grand-total row):")
    for label, mine, src_tot in [
        ("Units", len(units), 83),
        ("Sq Ft", sum(u.sqft or 0 for u in units), 71567),
        ("Rent (in-place)", sum(u.contract_rent for u in units), 113985),
    ]:
        flag = "OK" if abs(mine - src_tot) < 0.5 else "DIFF"
        print(f"    {label:16s} mine={mine:>12,.2f}  source={src_tot:>12,.2f}  [{flag}]")
    return 0


if __name__ == "__main__":
    sys.exit(main())
