"""Intake for the 'Rent Roll Analysis' export (single property per file).

Used for Delsan Court and Hertel Homes (same format; column positions differ
slightly between exports, so columns are located by their row-5 header text).

Layout: title rows 1-3 ("Property: <name>", "As of MM/DD/YY"), a header at row
5, a property banner at row 6, then one row per unit, closed by a "Totals for
<name>" block and a "Report Summary" block (both skipped / used to reconcile).

Columns (by header): Tenant Name (data sits ONE column right of the header),
Unit, Unit Type, Sq Ft, Market Rent, Rent, Vacancy Loss, Misc Charges,
Total Charges, Security Deposit, Move In, Move Out, Lease End.

  * Unit Type "N Bed M Bath" / "N Bed/M Bath [FP]" -> N BD / M BA; the raw string
    is kept as the floor plan so the "FP" layout stays distinct.
  * Vacant = tenant "<VACANT>" (its Rent column echoes market; zeroed here).
    In-place/contract rent = Rent for occupied, 0 for vacant -> the occupied
    total matches the source's "Occupied Unit Rent".
  * Market Rent column present -> market_from_inplace off.
  * "Misc Charges" -> Other Income (labeled verbatim per #32).
  * No lease-sign date -> Move In is the lease-start placeholder.
"""
import os
import re
import sys
from datetime import date, datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from rent_roll_processor import Unit, RollConfig, build_exhibits  # noqa: E402
from rent_roll_processor.naming import output_filename            # noqa: E402

BDBA = re.compile(r"(\d+)\s*Bed[\s/]*(\d+(?:\.\d+)?)\s*Bath", re.I)


def decode(ut):
    raw = str(ut or "").strip()
    m = BDBA.search(raw)
    if m:
        return raw, int(m.group(1)), float(m.group(2))
    return raw, None, None


def num(v):
    return v if isinstance(v, (int, float)) else None


def _money(v):
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        try:
            return float(v.replace(",", "").replace("$", "").strip())
        except ValueError:
            return None
    return None


def _parse_meta(ws):
    prop, asof = None, None
    for r in range(1, 5):
        v = ws.cell(r, 1).value
        if isinstance(v, str):
            if v.lower().startswith("property:"):
                prop = v.split(":", 1)[1].strip()
            m = re.search(r"as of\s+(\d{1,2})/(\d{1,2})/(\d{2,4})", v, re.I)
            if m:
                mm, dd, yy = (int(x) for x in m.groups())
                asof = date(2000 + yy if yy < 100 else yy, mm, dd)
    return prop, asof


def load(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True)["Sheet1"]
    prop, asof = _parse_meta(ws)

    H = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(5, c).value
        if isinstance(v, str):
            H[re.sub(r"\s+", " ", v).strip()] = c
    c_ten = H["Tenant Name"] + 1          # data is one column right of the header
    c_unit, c_ut = H["Unit"], H["Unit Type"]
    c_sf, c_mk, c_rent = H["Sq Ft"], H["Market Rent"], H["Rent"]
    c_vl, c_misc = H["Vacancy Loss"], H["Misc Charges"]
    c_mi = H.get("Move In")
    c_le = H.get("Lease End")
    c_mo = H.get("Move Out")

    def datecell(r, c):
        v = ws.cell(r, c).value if c else None
        return v if isinstance(v, datetime) else None

    units = []
    for r in range(7, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and a.strip().lower().startswith("totals for"):
            break                                   # end of unit rows
        ut = ws.cell(r, c_ut).value
        if not (isinstance(ut, str) and "bed" in ut.lower()):
            continue                                # skip banners / blanks
        fp, beds, baths = decode(ut)
        ten = str(ws.cell(r, c_ten).value or "").strip()
        vl = num(ws.cell(r, c_vl).value) or 0
        vacant = "vacant" in ten.lower() or vl > 0
        rent = num(ws.cell(r, c_rent).value) or 0
        misc = num(ws.cell(r, c_misc).value) or 0
        units.append(Unit.from_dict({
            "unit_id": str(ws.cell(r, c_unit).value or "").strip(),
            "unit_type": fp,                        # strict comp key
            "floor_plan": fp,
            "beds": beds,
            "baths": baths,
            "sqft": num(ws.cell(r, c_sf).value),
            "tenant_name": "Vacant" if vacant else (ten or None),
            "occupancy": "Vacant" if vacant else "Occupied",
            "market_rent": num(ws.cell(r, c_mk).value) or 0,
            "contract_rent": 0 if vacant else rent,
            "other_income": 0 if vacant else misc,
            "other_income_items": ({"Misc Charges": misc} if (misc and not vacant) else {}),
            "move_in": datecell(r, c_mi),
            "lease_start": datecell(r, c_mi),        # no sign date -> Move In
            "lease_end": datecell(r, c_le),
            "move_out": datecell(r, c_mo),
        }))

    # source Report Summary (label in col A, value in the 'Value'/rightmost col)
    summ = {}
    for r in range(7, ws.max_row + 1):
        lab = ws.cell(r, 1).value
        if isinstance(lab, str) and lab.strip() in (
                "# of Units", "Vacant Units", "Total Possible Rent",
                "Vacancy Rent", "Occupied Unit Rent"):
            for c in range(2, ws.max_column + 1):
                v = ws.cell(r, c).value
                if v not in (None, ""):
                    summ[lab.strip()] = _money(v)
                    break
    return units, prop, asof, summ


def main():
    args = sys.argv[1:]
    src = args[0] if args else None
    if not src:
        print("usage: intake_rentroll_analysis.py <source.xlsx> [out.xlsx]")
        return 2
    units, prop, asof, summ = load(src)
    prop = prop or "Property"
    out = args[1] if len(args) > 1 else output_filename(prop, asof)
    config = RollConfig(property_name=prop, as_of_date=asof, uw_market_rents={})
    build_exhibits(units, config, out, source_path=src, source_sheet="Sheet1")

    from rent_roll_processor.aggregate import is_occupied
    occ = sum(1 for u in units if is_occupied(u, config.model_occupied))
    vac = sum(1 for u in units if u.occupancy == "Vac")
    occ_rent = sum(u.contract_rent or 0 for u in units)
    zero = [u.unit_id for u in units if u.occupancy == "Occ" and not (u.contract_rent or 0)]
    print(f"Wrote {out}")
    print(f"  property={prop!r}  as_of={asof}")
    print(f"  units={len(units)}  occupied={occ}  vacant={vac}  occ%={occ/len(units):.2%}")
    if zero:
        print(f"  occupied $0-rent units -> market placeholder: {zero}")
    print("  reconcile (mine vs source Report Summary):")
    checks = [
        ("# of Units", len(units), summ.get("# of Units")),
        ("Vacant Units", vac, summ.get("Vacant Units")),
        ("Occupied Unit Rent", occ_rent, summ.get("Occupied Unit Rent")),
    ]
    for label, mine, src_tot in checks:
        if src_tot is None:
            print(f"    {label:20s} mine={mine:>12,.2f}  source=   (n/a)")
            continue
        flag = "OK" if abs(mine - src_tot) < 0.5 else "DIFF"
        print(f"    {label:20s} mine={mine:>12,.2f}  source={src_tot:>12,.2f}  [{flag}]")
    return 0


if __name__ == "__main__":
    sys.exit(main())
