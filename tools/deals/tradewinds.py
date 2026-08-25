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
    ),
    # market / move-in / move-out are not in this export -> left blank
    COLMAP=dict(unit=None, unittype=None, sqft=None, tenant=None,
                market=None, rent=None, other=None,
                lease_from=None, lease_to=None, movein=None, moveout=None),
    BD_BA_COL=None,          # no BD/BA in source -> floor-plan BD/BA left blank (client to fill)
    VACANT_TOKENS=("Vacant",),
    FLOOR_PLANS=None,        # auto-derive labels from distinct unit types
    PROPERTY="TradeWinds",
    ADDRESS="500 US 9",
    CITY="Waretown, NJ 08758",
    ASOF=date(2026, 7, 28),
)
