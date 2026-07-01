"""Intake for Colonial Pointe (Q1 2026) — an image-only (scanned) PDF rent roll.

The source has no machine-readable text (each page is a scanned image), so the
rent-roll rows on pages 2 (building 556) and 4 (building 558) were transcribed
here.  The two buildings are one property, combined into a single deal; unit ids
are prefixed with the building number to keep them unique.

Source columns: Unit | Type | Sq. Feet | Residents | Status | Lease End | Total
  * only ONE rent figure ("Total") -> used as both market and in-place rent
    (occupied); vacant carries 0 in-place.
  * no bed/bath, no move-in/lease-start, no other income/concession columns.
    -> BD/BA unknown (fill the OneLineRR Checking table); Recent Leases is
       empty (no lease-start data).
Flags: 556-000 "Leasing Mgr Storage Unit" and 558-307 "Maintenance Super Apt".
"""
import os
import re
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from rent_roll_processor import Unit, RollConfig, build_exhibits  # noqa: E402
from rent_roll_processor.naming import output_filename            # noqa: E402
from rent_roll_processor.schema import _coerce_date               # noqa: E402
from rent_roll_processor.mapping import (                         # noqa: E402
    write_unit_type_mapping, read_unit_type_mapping)

PROPERTY = "Colonial Pointe"
AS_OF = date(2026, 3, 31)   # Q1 2026 (no explicit as-of on the rent-roll pages)
# Notes that flag a non-residential / non-revenue unit -> Admin.
ADMIN_NOTE_RE = re.compile(r"storage|super|mainten|leasing|office|admin|employee|model", re.I)

# building, unit, type, sqft, residents, status, lease_end, total, note
ROWS = [
    # ---- Colonial Pointe - 556 (page 2) ----
    ("556", "000", "A-1", 845, "Cathleen Foster", "C", "12/31/26", 212, "Leasing Mgr Storage Unit"),
    ("556", "201", "A-1", 845, "Anna Zilenziger, Rodman Zilenziger", "C", "04/30/26", 3085, ""),
    ("556", "202", "J-2", 1209, "Laura Fleming, Robert McClellan", "C", "08/31/27", 3660, ""),
    ("556", "203", "E-2", 1146, "Janet Kost", "C", "11/30/26", 3565, ""),
    ("556", "204", "E-2", 1146, "Arlene Camporeale, Ignazio Camporeale", "C", "06/30/27", 3627, ""),
    ("556", "205", "F-2", 1141, "Elizabeth Cooper, Samuel Cooper", "C", "09/30/26", 3222, ""),
    ("556", "206", "E-2", 1146, "Josephine Makoujy", "C", "10/31/26", 3500, ""),
    ("556", "207", "C-1D", 1087, "Barbara Brown", "C", "07/31/26", 2950, ""),
    ("556", "208", "K-2", 1134, "Joy Eliezer", "C", "04/30/27", 3500, ""),
    ("556", "209", "K-2", 1134, "Jennifer Wiedemann, Kyle Wiedemann", "C", "02/28/27", 3408, ""),
    ("556", "210", "E-2", 1146, "Constance Franks", "C", "07/31/26", 3508, ""),
    ("556", "211", "F-2", 1141, "Robert Lynch, Gail Lynch", "C", "11/30/26", 3427, ""),
    ("556", "212", "E-2", 1146, "Karen Tileston", "C", "03/31/28", 3574, ""),
    ("556", "213", "H-2D", 1371, "Adele Brett", "C", "09/30/26", 3982, ""),
    ("556", "214", "F-2", 1141, "Michael Rizzotto, Pat Rizzotto", "C", "07/31/26", 3500, ""),
    ("556", "215", "D-1D", 980, "Barbara Mino", "C", "06/30/26", 3082, ""),
    ("556", "216", "K-2", 1134, "Richard Ege", "C", "01/31/27", 3518, ""),
    ("556", "217", "K-2", 1134, "Rozlyne Tessler", "C", "07/31/26", 3629, ""),
    ("556", "218", "E-2", 1146, "Kersti Kimler", "C", "01/31/27", 3582, ""),
    ("556", "219", "A-1", 845, "Frederick Schulze", "C", "02/28/27", 2850, ""),
    ("556", "220", "A-1", 845, "Audrey Roberts", "MTM", "05/31/26", 2822, ""),
    ("556", "301", "A-1", 845, "Brianne O' Rourke", "C", "06/30/26", 2800, ""),
    ("556", "302", "J-2", 1209, "Rita Grogan", "C", "04/30/26", 3814, ""),
    ("556", "303", "E-2", 1146, "Charles Grenner", "C", "10/31/26", 3543, ""),
    ("556", "304", "E-2", 1146, "Rhonda Golub, Mark Frohman", "C", "07/31/26", 3475, ""),
    ("556", "305", "F-2", 1141, "Peter Macaluso, MaryAnn Macaluso", "C", "04/30/26", 3441, ""),
    ("556", "306", "E-2", 1146, "Christopher Nixon", "C", "07/31/26", 3425, ""),
    ("556", "307", "B-1D", 980, "Brenda McElkenny", "C", "06/30/26", 2939, ""),
    ("556", "308", "K-2", 1134, "Adrienne Messina, Nicholas Messina", "C", "05/31/26", 3475, ""),
    ("556", "309", "K-2", 1134, "Jane Barovick", "C", "05/31/26", 3569, ""),
    ("556", "310", "E-2", 1146, "John Rosica", "C", "01/31/27", 3533, ""),
    ("556", "311", "F-2", 1141, "Clarie Regen", "C", "07/31/26", 3450, ""),
    ("556", "312", "E-2", 1146, "Michael Liszczynsky", "C", "09/30/27", 3738, ""),
    ("556", "313", "H-2D", 1371, "Michael Rosenberg", "C", "12/31/26", 4238, ""),
    ("556", "314", "G-2D", 1434, "Leonard Birkhahn, Iris Birkhahn", "C", "02/28/27", 3900, ""),
    ("556", "315", "F-2", 1141, "Geraldine Caldiero", "C", "01/31/27", 3532, ""),
    ("556", "316", "B-1D", 980, "Cecilia Turula", "C", "11/30/26", 2872, ""),
    ("556", "317", "K-2", 1134, "Lois Bodner", "C", "09/30/27", 3533, ""),
    ("556", "318", "K-2", 1134, "Angelo Ingrassia", "C", "11/30/26", 3458, ""),
    ("556", "319", "E-2", 1146, "Francis DeBlock", "C", "05/31/26", 3482, ""),
    ("556", "320", "A-1", 845, "George Schluger", "C", "07/31/27", 2914, ""),
    ("556", "321", "A-1", 845, "Takuma Oda", "C", "03/31/27", 2848, ""),
    ("556", "322", "A-1", 845, "Mark Abidargham, Sally Khalil", "C", "07/31/26", 2823, ""),
    ("556", "401", "A-1", 845, "Dana Stamler", "C", "04/30/26", 2800, ""),
    ("556", "402", "J-2", 1209, "Dina Messery, Craig Messery", "C", "06/30/26", 3634, ""),
    ("556", "403", "E-2", 1146, "Matthew Brizzi, Sara Brizzi", "C", "10/31/26", 3610, ""),
    ("556", "404", "E-2", 1146, "Donna Mancinelli", "C", "08/31/27", 3597, ""),
    ("556", "405", "F-2", 1141, "Dana Sopel", "C", "10/31/26", 3629, ""),
    ("556", "406", "E-2", 1146, "Hilda Jorge-Garcia, Fermin Garcia", "C", "07/31/26", 3425, ""),
    ("556", "407", "B-1D", 980, "Thomas Tashjian", "C", "05/31/26", 2925, ""),
    ("556", "408", "K-2", 1134, "Camille Samuel, Julyet Samuel, Patricia Samuel", "C", "11/30/26", 3533, ""),
    ("556", "409", "K-2", 1134, "Marissa Liza", "C", "04/30/26", 3577, ""),
    ("556", "410", "E-2", 1146, "Sally Kady, John Kady", "C", "09/30/26", 3565, ""),
    ("556", "411", "F-2", 1141, "Gail Marquard", "C", "10/31/27", 3467, ""),
    ("556", "412", "E-2", 1146, "Aga Montes, Vin Montes", "C", "09/30/26", 3500, ""),
    ("556", "413", "H1-2D", 1505, "Anthony Ruggiero", "C", "09/30/26", 4065, ""),
    ("556", "414", "G-2D", 1434, "Joseph Geronimo, Anna Geronimo", "C", "04/30/27", 3964, ""),
    ("556", "415", "F-2", 1141, "Janet Koeller", "C", "05/31/26", 3443, ""),
    ("556", "416", "B-1D", 980, "Sarah Kuldip", "C", "02/28/27", 3028, ""),
    ("556", "417", "K-2", 1134, "Alicia Marcucilli", "C", "02/28/27", 3593, ""),
    ("556", "418", "K-2", 1134, "Ani Gabrellian", "C", "08/31/26", 3745, ""),
    ("556", "419", "E-2", 1146, "David Smiley", "C", "05/31/26", 3425, ""),
    ("556", "420", "A-1", 845, "Vacant Unit", "", "", 2900, ""),
    ("556", "421", "A-1", 845, "Carol Lobban", "C", "09/30/26", 2848, ""),
    ("556", "422", "A-1", 845, "Kristy Rothfritz", "C", "12/31/26", 2828, ""),
    # ---- Colonial Pointe - 558 (page 4) ----
    ("558", "101", "P1-1 AFF L", 710, "Romualdo Turelli", "C", "03/31/27", 977, ""),
    ("558", "102", "L-2 AFF VL", 881, "Lori Clolinger", "C", "03/31/27", 625, ""),
    ("558", "103", "L-2 AFF VL", 881, "Amanda Parisi", "C", "03/31/27", 704, ""),
    ("558", "104", "N-2 AFF VL", 960, "Jacqueline Craven", "C", "12/31/26", 623, ""),
    ("558", "105", "N-2 AFF L", 960, "Amelia DeLaRosa", "C", "04/30/26", 1073, ""),
    ("558", "106", "M-3 AFF VL", 1203, "Jahnita Merced", "C", "03/31/27", 887, ""),
    ("558", "107", "M-2 AFF L", 1047, "Marilyn Betancourt", "C", "04/30/26", 1073, ""),
    ("558", "201", "P-1 AFF L", 984, "Demetria Moore", "C", "02/28/27", 1133, ""),
    ("558", "202", "L-2 AFF M", 881, "Donna Mango, Peter Mango", "C", "02/28/27", 1386, ""),
    ("558", "203", "L-2 AFF M", 881, "Ingrid Daniels, LaNique Watts", "C", "02/28/27", 1465, ""),
    ("558", "204", "N-2 AFF M", 960, "Karen Williamson", "C", "08/31/26", 1297, ""),
    ("558", "205", "N-2 AFF M", 960, "Veronica Candia-Toledo", "MTM", "", 1313, ""),
    ("558", "206", "M-3 AFF L", 1203, "Vacant", "", "", 1650, ""),
    ("558", "207", "M-3 AFF M", 1203, "Wanda DeLeon", "C", "12/31/26", 1644, ""),
    ("558", "208", "Q-1 AFF M", 705, "Eva Noa", "C", "10/31/26", 1123, ""),
    ("558", "301", "P-1", 984, "Daniel Mazza", "C", "08/31/27", 2710, ""),
    ("558", "302", "L-2", 881, "Harold Schultz", "C", "06/30/26", 2650, ""),
    ("558", "303", "L-2", 881, "Kamelia Rafeh-Saheli", "C", "05/31/26", 2631, ""),
    ("558", "304", "N-2 AFF M", 960, "Darlene Rivetti, Michael Rivetti", "C", "01/31/27", 1385, ""),
    ("558", "305", "N-2 AFF M", 960, "Tanya Cintron", "C", "05/31/26", 1335, ""),
    ("558", "306", "M-3 AFF M", 1203, "Jose Ortiz, Priscilla Rosario", "C", "12/31/26", 1514, ""),
    ("558", "307", "M-3", 1203, "Jose Ucros", "C", "03/31/27", 2000, "Maintenaince Super Apt"),
    ("558", "308", "Q-1", 705, "Jacklyn Marshall", "C", "03/31/27", 2300, ""),
]


def build_units():
    units = []
    for bldg, unit, utype, sqft, residents, status, lease_end, total, note in ROWS:
        st = str(status).strip().upper()
        name = residents.strip()
        vacant = "vacant" in name.lower()
        if note and ADMIN_NOTE_RE.search(note):
            occ = "Admin"          # non-revenue (storage / super / etc.)
        elif vacant:
            occ = "Vacant"
        else:
            occ = "Occupied"
        units.append(Unit.from_dict({
            "unit_id": f"{bldg}-{unit}",
            "unit_type": utype,
            "floor_plan": utype,        # BD/BA unknown -> floor plan = type code
            "beds": None,
            "baths": None,
            "sqft": sqft,
            "tenant_name": name,
            "occupancy": occ,
            "market_rent": total,                       # only one rent figure in source
            "contract_rent": 0 if vacant else total,    # in-place (0 if vacant)
            "mtm": "MTM" if st == "MTM" else None,
            "lease_end": _coerce_date(lease_end),
            "designation": note or None,
        }))
    return units


def main():
    args = sys.argv[1:]
    units = build_units()

    # Step 1: emit the fill-in Unit Type Mapping table, then stop.
    if args and args[0] == "map":
        path = args[1] if len(args) > 1 else f"Unit Type Mapping - {PROPERTY}.xlsx"
        write_unit_type_mapping(units, path, PROPERTY)
        print(f"Wrote mapping table {path} "
              f"({len({u.unit_type for u in units})} unit types)")
        return 0

    # Step 2 (optional): incorporate a filled mapping table.
    out = args[0] if args else output_filename(PROPERTY)
    utmap = read_unit_type_mapping(args[1]) if len(args) > 1 else {}
    config = RollConfig(property_name=PROPERTY, as_of_date=AS_OF,
                        uw_market_rents={}, unit_type_map=utmap)
    build_exhibits(units, config, out)

    from rent_roll_processor.aggregate import is_occupied
    occ = sum(1 for u in units if is_occupied(u, config.model_occupied))
    vac = sum(1 for u in units if u.occupancy == "Vac")
    adm = sum(1 for u in units if u.occupancy == "Admin")
    print(f"Wrote {out}")
    print(f"  units={len(units)}  occupied={occ}  vacant={vac}  admin(non-rev)={adm}  "
          f"occ%={occ/len(units):.2%}  mapping={'applied' if utmap else 'none (types as floor plans)'}")
    # reconcile rent totals per building against the source page totals
    for bldg, src_total in (("556", 218106), ("558", 33500)):
        tot = sum(u.market_rent for u in units if u.unit_id.startswith(bldg + "-"))
        n = sum(1 for u in units if u.unit_id.startswith(bldg + "-"))
        print(f"  {bldg}: {n} units, Total sum={tot:,.0f} (source {src_total:,})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
