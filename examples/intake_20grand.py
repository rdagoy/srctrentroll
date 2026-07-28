"""Intake for 20 Grand Ave (Yardi 'Rent Roll with Lease Charges').

Sheet 'Report1'.  A charge-ledger layout (like Maven): one main row per unit
carrying the market rent + first charge code, then one sub-row per additional
charge code, closed by a "Total" row.  Sections are marked by banner rows:

  * "Current/Notice/Vacant Residents"  <- the as-of snapshot (parsed)
  * "Future Residents/Applicants"      <- upcoming leases for units already
                                          counted; SKIPPED (would double-count)
  * "Summary Groups" / charge summary  <- SKIPPED

Mixed use: 2 commercial units (15SDCOMM) + residential 201..526.

Columns: A Unit | B Unit Type | C Unit Sq Ft | D Resident | E Name
  | F Market Rent | G Charge Code | H Amount | I Resident Deposit
  | J Other Deposit | K Move In | L Lease Expiration | M Move Out | N Balance

Charge-code mapping (validated against the source's current-only charge
summary, total 327,746.45):
  * BASE (contract rent) = Rent, Retail, ComRent
  * CONCESSIONS (negative) = AmorConc, Amenconc, ParkConc, MIConc, RenConc,
      CommConc, Super   (Super = superintendent credit, netted per analyst)
  * OTHER INCOME (everything else, itemized) = Retax, CAM, Amenity, Trash,
      watsew, Parking, PetRent, MTM
No lease-sign date in the source -> Move In is the lease-start placeholder.
"""
import os
import re
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from rent_roll_processor import Unit, RollConfig, build_exhibits  # noqa: E402
from rent_roll_processor.naming import output_filename            # noqa: E402

PROPERTY = "20 Grand Ave"
AS_OF = date(2026, 7, 30)          # source header "As Of = 07/30/2026"

BASE = {"Rent", "Retail", "ComRent"}
CONC = {"AmorConc", "Amenconc", "ParkConc", "MIConc", "RenConc", "CommConc", "Super"}
BANNERS = {"Current/Notice/Vacant Residents", "Future Residents/Applicants",
           "Summary Groups"}
BDBA_RE = re.compile(r"(\d)B(\d)")


def decode(utype):
    """15SD<plan> -> (floor_plan, beds, baths).  Plan variants B/D (likely a den
    layout) keep distinct unit_type codes but share a BD/BA floor plan."""
    rest = str(utype or "")[4:]
    if rest == "Stu":
        return "0 BD / 1 BA", 0, 1          # studio
    if rest == "COMM":
        return "Commercial", None, None     # commercial: no BD/BA
    m = BDBA_RE.match(rest)
    if m:
        b, ba = int(m.group(1)), int(m.group(2))
        return f"{b} BD / {ba} BA", b, ba
    return str(utype or "").strip(), None, None


def num(v):
    return v if isinstance(v, (int, float)) else None


def load(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True)["Report1"]

    def is_banner(r):
        a = ws.cell(r, 1).value
        return isinstance(a, str) and a.strip() in BANNERS

    def is_unit_start(r):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        return (isinstance(a, str) and a.strip() and not is_banner(r)
                and isinstance(b, str) and b.strip())

    stops = sorted({r for r in range(5, ws.max_row + 1)
                    if is_banner(r) or is_unit_start(r)} | {ws.max_row + 1})

    units, section = [], None
    for r in range(5, ws.max_row + 1):
        if is_banner(r):
            section = ws.cell(r, 1).value.strip()
            continue
        if section != "Current/Notice/Vacant Residents" or not is_unit_start(r):
            continue
        end = min(x for x in stops if x > r)          # next unit-start or banner
        charges = {}
        for rr in range(r, end):
            g, h = ws.cell(rr, 7).value, ws.cell(rr, 8).value
            if isinstance(g, str) and g not in ("Charge", "Total") and isinstance(h, (int, float)):
                charges[g] = charges.get(g, 0) + h

        fp, beds, baths = decode(ws.cell(r, 2).value)
        name = str(ws.cell(r, 5).value or "").strip()
        resid = str(ws.cell(r, 4).value or "").strip()
        vacant = name.upper() == "VACANT" or resid.upper() == "VACANT"

        contract = sum(v for k, v in charges.items() if k in BASE)
        concession = sum(v for k, v in charges.items() if k in CONC)
        oi_items = {k: v for k, v in charges.items() if k not in BASE and k not in CONC}

        units.append(Unit.from_dict({
            "unit_id": str(ws.cell(r, 1).value).strip(),
            "unit_type": str(ws.cell(r, 2).value).strip(),
            "floor_plan": fp,
            "beds": beds,
            "baths": baths,
            "sqft": num(ws.cell(r, 3).value),
            "tenant_name": name or ("Vacant" if vacant else None),
            "occupancy": "Vacant" if vacant else "Occupied",
            "market_rent": num(ws.cell(r, 6).value) or 0,        # F
            "contract_rent": 0 if vacant else contract,          # sum(BASE)
            "concession": 0 if vacant else concession,           # sum(CONC)
            "other_income": sum(oi_items.values()),
            "other_income_items": oi_items,
            "move_in": ws.cell(r, 11).value,
            "lease_start": ws.cell(r, 11).value,   # no sign date -> Move In placeholder
            "lease_end": ws.cell(r, 12).value,     # Lease Expiration
            "move_out": ws.cell(r, 13).value,
        }))
    return units


def main():
    args = sys.argv[1:]
    src = args[0] if args else None
    out = args[1] if len(args) > 1 else output_filename(PROPERTY)
    if not src:
        print("usage: intake_20grand.py <source.xlsx> [out.xlsx]")
        return 2
    units = load(src)
    config = RollConfig(property_name=PROPERTY, as_of_date=AS_OF, uw_market_rents={})
    build_exhibits(units, config, out)

    from rent_roll_processor.aggregate import is_occupied
    occ = sum(1 for u in units if is_occupied(u, config.model_occupied))
    vac = sum(1 for u in units if u.occupancy == "Vac")
    print(f"Wrote {out}")
    print(f"  units={len(units)}  occupied={occ}  vacant={vac}  occ%={occ/len(units):.2%}")
    print("  reconcile (mine vs source Summary/charge totals):")
    contract = sum(u.contract_rent for u in units)
    conc = sum(u.concession for u in units)
    other = sum(u.other_income for u in units)
    for label, mine, src_tot in [
        ("Units", len(units), 98),
        ("Occupied units", occ, 86),
        ("Vacant units", vac, 12),
        ("Sq Ft", sum(u.sqft or 0 for u in units), 98297.09),
        ("Market Rent", sum(u.market_rent for u in units), 364604.31),
        ("Base rent", contract, 323621.68),
        ("Other income", other, 27549.51),
        ("Concessions", conc, -23424.74),
        ("Lease charges", contract + other + conc, 327746.45),
    ]:
        flag = "OK" if abs(mine - src_tot) < 0.5 else "DIFF"
        print(f"    {label:16s} mine={mine:>13,.2f}  source={src_tot:>13,.2f}  [{flag}]")
    return 0


if __name__ == "__main__":
    sys.exit(main())
