"""Intake / normalization for the Maven @ 806 (29SC / RealPage) rent roll.

Source layout (sheet 'Detailed'):
    A3 property, A4 as-of date, header row 7, unit blocks from row 8.
    Each unit = a main row (Bldg-Unit + unit details + first charge) followed by
    charge-code sub-rows (blank Bldg-Unit) and a 'Charge Total:' row.

    Columns: A Bldg-Unit | B Unit Type | C SQFT | D Unit Status | E Resident |
    F Move-In | G Lease Start | H Lease End | I Expected Move-Out |
    J Advertised Rent | K Ledger | L Charge Code | ... | O Scheduled Charges | ...

Mapping decisions (flagged to the analyst):
  * market_rent   = Advertised Rent (col J).
  * contract_rent = sum of every 'Rent-*' charge code (base + premiums +
    negative amenity) — the effective in-place rent.
  * concession    = sum of 'Concession-*' codes (already negative; e.g. the
    employee rent discount). Nets into in-place rent per the model.
  * other_income  = all remaining charges (Building Protection, Pet, Pest,
    Trash, Water, WiFi/Cable, RUBS, MTM Fee).
  * Unit type NxM  -> N bed / M bath (e.g. 1x1B_C -> 1 BD / 1 BA). No reno tier.
  * Occupancy: 'vacant' -> Vac; 'notice'/'occupied' -> Occ (NTV counts occupied).
"""
import os
import re
import sys
from datetime import date

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from rent_roll_processor import Unit, RollConfig, build_exhibits  # noqa: E402
from rent_roll_processor.naming import output_filename  # noqa: E402

PROPERTY = "Maven @ 806"
AS_OF = date(2026, 5, 15)
SRC = sys.argv[1] if len(sys.argv) > 1 else None
OUT = sys.argv[2] if len(sys.argv) > 2 else output_filename(PROPERTY)

TYPE_RE = re.compile(r"^(\d+)x(\d+)")
# A unit header row is identified by its Unit Type cell (e.g. "1x1B_C",
# "2x1A_C") — robust to the two Bldg-Unit id formats in this report
# ("800-1A" and "CL - 806-1").  Charge sub-rows have a blank Unit Type.
TYPE_HDR_RE = re.compile(r"^\s*\d+x\d+")


def decode(code):
    """1x1B_C -> ('1 BD / 1 BA', 1, 1, renovated?). No reno tier in this file."""
    raw = str(code).strip()
    m = TYPE_RE.match(raw)
    beds = baths = None
    if m:
        beds, baths = int(m.group(1)), int(m.group(2))
    reno = bool(re.search(r"_R\b|R$", raw))   # future-proof; none present here
    fp = f"{beds} BD / {baths} BA" if beds is not None else raw
    return fp, beds, baths, reno, raw


def occ_from_status(status):
    s = str(status or "").lower()
    if "vacant" in s:
        return "Vacant"
    return "Occupied"          # occupied / notice (NTV) both -> occupied


def is_unit_row(unit_type_val):
    """True for a unit header row, keyed off the Unit Type cell (col B)."""
    return bool(unit_type_val and TYPE_HDR_RE.match(str(unit_type_val)))


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Detailed"]
    units = []
    cur = None

    def finalize(c):
        if c:
            units.append(Unit.from_dict(c["fields"] | {
                "contract_rent": c["rent"],
                "concession": c["conc"],
                "other_income": sum(c["oi_items"].values()),
                "other_income_items": dict(c["oi_items"]),
            }))

    for r in range(8, ws.max_row + 1):
        a = ws.cell(r, 1).value
        # Stop at the 'Future Resident Details' section (empty here).
        if isinstance(a, str) and a.strip() == "Future Resident Details":
            break
        if is_unit_row(ws.cell(r, 2).value):
            finalize(cur)
            status = ws.cell(r, 4).value
            occ = occ_from_status(status)
            fp, beds, baths, reno, raw = decode(ws.cell(r, 2).value)
            name = str(ws.cell(r, 5).value or "").strip()
            from collections import defaultdict
            cur = {"rent": 0.0, "conc": 0.0, "oi_items": defaultdict(float), "fields": {
                "unit_id": a.strip(),
                "unit_type": raw,
                "floor_plan": fp,
                "beds": beds,
                "baths": baths,
                "renovated": "Yes" if reno else None,
                "reno_type": "Full" if reno else None,
                "sqft": ws.cell(r, 3).value,
                "tenant_name": name or ("VACANT" if occ == "Vacant" else None),
                "occupancy": occ,
                "market_rent": ws.cell(r, 10).value,      # Advertised Rent
                "move_in": ws.cell(r, 6).value,
                "lease_start": ws.cell(r, 7).value,
                "lease_end": ws.cell(r, 8).value,
                "move_out": ws.cell(r, 9).value,   # Expected Move-Out
            }}
        # Accumulate charge-code line (present on main row and sub-rows).
        code = ws.cell(r, 12).value
        amt = ws.cell(r, 15).value
        if (cur and code and str(code).strip() not in ("Charge Total:", "-")
                and isinstance(amt, (int, float))):
            c = str(code)
            if c.startswith("Rent-"):
                cur["rent"] += amt
            elif c.startswith("Concession"):
                cur["conc"] += amt
            else:
                cur["oi_items"][c] += amt   # keep each other-income line item
    finalize(cur)
    return units


def main():
    if not SRC:
        print("usage: intake_maven.py <source.xlsx> [out.xlsx]")
        return 2
    units = load(SRC)
    config = RollConfig(
        property_name=PROPERTY,
        as_of_date=AS_OF,
        uw_market_rents={},        # LTL left blank
    )
    build_exhibits(units, config, OUT)

    from rent_roll_processor.aggregate import is_occupied
    occ = sum(1 for u in units if is_occupied(u, config.model_occupied))
    vac = sum(1 for u in units if u.occupancy == "Vac")
    print(f"Wrote {OUT}")
    print(f"  units={len(units)}  occupied={occ}  vacant={vac}  occ%={occ/len(units):.2%}")
    from collections import Counter
    fp = Counter(u.floor_plan for u in units)
    for k in sorted(fp):
        print(f"    {k}: {fp[k]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
