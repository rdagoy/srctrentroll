"""Per-deal config: TradeWinds (Waretown, NJ) — itemized rent roll.

Source: Tradewinds_RR_Itemized_7.28.26.xlsx, tab 'RR Itemized'.
Itemized format: one unit block spans several charge rows ending in 'Net:'.
Charges are bucketed -> Rent (AA), Section 8 -> Subsidy (AB), all other
residential charges -> Other Income (AD). Commercial and $0 non-revenue
units are excluded from the residential roll.
"""
from datetime import date

CONFIG = dict(
    FORMAT="itemized",
    SOURCE_SHEET="RR Itemized",
    ITEMIZED=dict(
        COLS=dict(unit="A", name="B", utype="C", sqft="D", lease="E",
                  charge="H", monthly="J"),
        HEADER_UNIT="Unit",
        RENT=["Rent"],
        SUBSIDY=["Section 8"],
        COMMERCIAL=["Commercial Rent", "CAM", "RET", "Insurance"],
        BLANK_TYPE_AS_UNIT=True,   # units with no type (Office) take their unit id as the type
    ),
    # market / move-in / move-out are not in this export -> left blank
    COLMAP=dict(unit=None, unittype=None, sqft=None, tenant=None,
                market=None, rent=None, other=None,
                lease_from=None, lease_to=None, movein=None, moveout=None),
    BD_BA_COL=None,          # no BD/BA in source
    VACANT_TOKENS=("Vacant",),
    # Client-provided BD/BA table (this deal only), in the client's order.
    FLOOR_PLANS=[dict(label=t, unittype=t, bd=bd, ba=ba,
                      affordable="No", renovated="No", reno_type="Classic")
                 for (t, bd, ba) in [
        ("S.2.B", 0, 1.0), ("S.2.C", 0, 1.0),
        ("2.2.F", 2, 2.0), ("2.2.G", 2, 2.0), ("2.2.I", 2, 2.0), ("2.2.J", 2, 2.0),
        ("2.2.H", 2, 2.0), ("2.2.K", 2, 2.0), ("2.2.D", 2, 2.0),
        ("3.2.A", 3, 2.0), ("3.2.E", 3, 2.0),
        ("Marina", 1, 1.0), ("Veranda", 2, 1.0), ("Santa Ana", 2, 1.0),
        ("Bayside", 2, 2.0), ("Oasis", 2, 2.0), ("Oceanic", 2, 2.0),
        ("Office", 1, 1.0),
    ]],
    PROPERTY="TradeWinds",
    ADDRESS="500 US 9",
    CITY="Waretown, NJ 08758",
    ASOF=date(2026, 7, 28),
)
