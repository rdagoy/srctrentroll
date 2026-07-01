"""Aggregation engine.

All summary math lives here so it can be unit-tested independently of Excel
writing.  Every formula below was reverse-engineered from, and validated
against, the Heritage Hill Estates reference exhibits file.

Conventions
-----------
* "occupied" = Occ or Model (see Unit.is_occupied).
* Per-unit "Market Rent" is the source column U (a per-unit market/asking
  figure).  The underwriting market rent (config.uw_market_rents) is separate
  and only feeds the LTL columns + the AE total.
* Rolling-lease windows count occupied units whose lease_start falls in
  [as_of - N days, as_of] and average their contract rent.
"""
from __future__ import annotations

import calendar
from dataclasses import dataclass
from datetime import date
from typing import List, Optional, Dict, Any

from .schema import Unit, RollConfig, OCC, MODEL

NA = "n/a"


def is_occupied(unit: Unit, model_occupied: bool = False) -> bool:
    """Occupancy used by every summary.  Model units count as occupied only
    when ``model_occupied`` is True (see RollConfig.model_occupied)."""
    if unit.occupancy == MODEL:
        return model_occupied
    return unit.occupancy == OCC


def in_place_rent(unit: Unit, net_concession: bool = False) -> float:
    """In-place (contract) rent used by the summaries.  When ``net_concession``
    is on, the concession (a negative number) is netted in, so e.g. a model
    unit whose rent = market and concession = -market contributes 0."""
    v = unit.contract_rent or 0.0
    if net_concession:
        v += unit.concession or 0.0
    return v
# Rolling lease windows.  Each "N days" column is really an N/30 calendar-month
# look-back (this matches the reference exhibits, e.g. a lease 61 days out still
# lands in the "60 Days" column because it is within 2 calendar months).
RECENT_WINDOWS = (180, 120, 90, 60, 30)
WINDOW_MONTHS = {180: 6, 120: 4, 90: 3, 60: 2, 30: 1}


def _avg(values: List[float]) -> Optional[float]:
    return sum(values) / len(values) if values else None


def _minus_months(d: date, months: int) -> date:
    """Subtract whole calendar months, clamping the day to month length."""
    m = d.month - 1 - months
    y = d.year + m // 12
    m = m % 12 + 1
    day = min(d.day, calendar.monthrange(y, m)[1])
    return date(y, m, day)


@dataclass
class GroupSummary:
    """One row of the Unit Mix / Bed Mix table (and the shared Recent Leases
    left-hand columns)."""
    key: str                 # floor plan name or bed label
    occ_units: int
    tot_units: int
    sqft: float              # total SF (all units)
    avg_sf: float
    mkt_rent: float          # sum of per-unit market rent (all units)
    avg_mkt_unit: float      # mkt_rent / tot_units
    avg_mkt_sf: float        # mkt_rent / sqft
    cont_rent: float         # sum of contract rent (all units)
    avg_cont_unit: Optional[float]   # cont_rent / occ_units
    avg_cont_sf: Optional[float]     # cont_rent / occ_sqft
    max_rent: float          # max per-unit market rent
    other_inc: float         # sum other income (all units)
    occ_sqft: float          # SF of occupied units
    # sort helpers
    beds: float = 0.0
    baths: float = 0.0


@dataclass
class RecentLeaseRow:
    key: str
    occ_units: int
    tot_units: int
    sqft: float
    avg_sf: float
    avg_mkt_unit: float           # J - same as unit-mix avg_mkt_unit
    occ_market_rent: float        # K - avg market rent of OCCUPIED units
    in_place_rent: Optional[float]  # L - avg contract of occupied
    pct_of_market: Optional[float]  # M - L / K
    windows: Dict[int, Any]       # N..W -> {"avg": float|"n/a", "count": int}
    beds: float = 0.0
    baths: float = 0.0


@dataclass
class Totals:
    occ_units: int
    tot_units: int
    sqft: float
    avg_sf: float
    mkt_rent: float
    avg_mkt_unit: float
    avg_mkt_sf: float
    cont_rent: float
    avg_cont_unit: Optional[float]
    avg_cont_sf: Optional[float]
    max_rent: float                # weighted avg of group max by tot_units
    other_inc: float
    occ_sqft: float
    # Recent-leases-specific totals
    occ_market_rent: float         # sum(market of occ)/occ_units
    in_place_rent: Optional[float]
    pct_of_market: Optional[float]
    window_counts: Dict[int, int]
    # Pres. RR / audit totals
    conc: float
    emp_disc: float
    uw_market_total: float         # sum of AE over all units
    ltl_total: float
    ltl_pct_total: Optional[float]


def _summarize_group(key: str, units: List[Unit], model_occupied: bool = False,
                     net_concession: bool = False) -> GroupSummary:
    occ = [u for u in units if is_occupied(u, model_occupied)]
    tot_units = len(units)
    occ_units = len(occ)
    sqft = sum(u.sqft for u in units)
    occ_sqft = sum(u.sqft for u in occ)
    mkt_rent = sum(u.market_rent for u in units)
    cont_rent = sum(in_place_rent(u, net_concession) for u in units)
    other_inc = sum(u.other_income for u in units)
    # "Max Rent" = the top in-place (contract) rent among occupied units.
    max_rent = max((in_place_rent(u, net_concession) for u in occ), default=0.0)
    beds = units[0].beds if units else 0.0
    baths = units[0].baths if units else 0.0
    return GroupSummary(
        key=key,
        occ_units=occ_units,
        tot_units=tot_units,
        sqft=sqft,
        avg_sf=sqft / tot_units if tot_units else 0.0,
        mkt_rent=mkt_rent,
        avg_mkt_unit=mkt_rent / tot_units if tot_units else 0.0,
        avg_mkt_sf=mkt_rent / sqft if sqft else 0.0,
        cont_rent=cont_rent,
        avg_cont_unit=cont_rent / occ_units if occ_units else None,
        avg_cont_sf=cont_rent / occ_sqft if occ_sqft else None,
        max_rent=max_rent,
        other_inc=other_inc,
        occ_sqft=occ_sqft,
        beds=beds,
        baths=baths,
    )


def _group_units(units: List[Unit], key_fn) -> "list[tuple[str, list[Unit]]]":
    """Group preserving first-seen order, then sort by (beds, baths, key)."""
    groups: Dict[Any, List[Unit]] = {}
    for u in units:
        groups.setdefault(key_fn(u), []).append(u)
    items = list(groups.items())
    items.sort(key=lambda kv: (kv[1][0].beds, kv[1][0].baths, str(kv[0])))
    return items


def unit_mix(units: List[Unit], model_occupied: bool = False,
             net_concession: bool = False) -> List[GroupSummary]:
    """Per floor-plan summary rows, sorted by beds/baths/plan."""
    return [_summarize_group(str(k), g, model_occupied, net_concession)
            for k, g in _group_units(units, lambda u: u.floor_plan)]


def bed_mix(units: List[Unit], model_occupied: bool = False,
            net_concession: bool = False) -> List[GroupSummary]:
    """Per bed-count summary rows."""
    out = []
    for k, g in _group_units(units, lambda u: u.beds):
        beds = g[0].beds
        label = int(beds) if float(beds).is_integer() else beds
        out.append(_summarize_group(str(label), g, model_occupied, net_concession))
    return out


def _window_stat(occ_units: List[Unit], as_of: date, days: int,
                 net_concession: bool = False) -> Dict[str, Any]:
    start = _minus_months(as_of, WINDOW_MONTHS[days])
    sel = [u for u in occ_units
           if u.lease_start and start <= u.lease_start <= as_of]
    if not sel:
        return {"avg": NA, "count": 0}
    return {"avg": sum(in_place_rent(u, net_concession) for u in sel) / len(sel),
            "count": len(sel)}


def recent_leases(units: List[Unit], as_of: date, model_occupied: bool = False,
                  net_concession: bool = False) -> List[RecentLeaseRow]:
    rows = []
    for k, g in _group_units(units, lambda u: u.floor_plan):
        occ = [u for u in g if is_occupied(u, model_occupied)]
        tot_units = len(g)
        occ_units = len(occ)
        sqft = sum(u.sqft for u in g)
        mkt_all = sum(u.market_rent for u in g)
        occ_market = _avg([u.market_rent for u in occ])
        in_place = _avg([in_place_rent(u, net_concession) for u in occ])
        rows.append(RecentLeaseRow(
            key=str(k),
            occ_units=occ_units,
            tot_units=tot_units,
            sqft=sqft,
            avg_sf=sqft / tot_units if tot_units else 0.0,
            avg_mkt_unit=mkt_all / tot_units if tot_units else 0.0,
            occ_market_rent=occ_market or 0.0,
            in_place_rent=in_place,
            pct_of_market=(in_place / occ_market) if (in_place and occ_market) else None,
            windows={d: _window_stat(occ, as_of, d, net_concession) for d in RECENT_WINDOWS},
            beds=g[0].beds,
            baths=g[0].baths,
        ))
    return rows


def compute_totals(units: List[Unit], config: RollConfig,
                   mix: Optional[List[GroupSummary]] = None) -> Totals:
    model_occupied = config.model_occupied
    net = config.net_concession
    mix = mix if mix is not None else unit_mix(units, model_occupied, net)
    occ = [u for u in units if is_occupied(u, model_occupied)]
    tot_units = len(units)
    occ_units = len(occ)
    sqft = sum(u.sqft for u in units)
    occ_sqft = sum(u.sqft for u in occ)
    mkt_rent = sum(u.market_rent for u in units)
    cont_rent = sum(in_place_rent(u, net) for u in units)
    other_inc = sum(u.other_income for u in units)
    conc = sum(u.concession for u in units)
    emp_disc = sum(u.emp_discount for u in units)

    occ_market = _avg([u.market_rent for u in occ]) or 0.0
    in_place = _avg([in_place_rent(u, net) for u in occ])

    # Max Rent total = weighted average of per-group max by total units.
    if tot_units and mix:
        max_rent = sum(g.max_rent * g.tot_units for g in mix) / tot_units
    else:
        max_rent = 0.0

    window_counts = {
        d: _window_stat(occ, config.as_of_date, d, net)["count"]
        for d in RECENT_WINDOWS
    }

    # Underwriting market + loss-to-lease.
    uw_total = 0.0
    ltl_total = 0.0
    for u in units:
        uw = config.uw_market(u.floor_plan)
        if uw is not None:
            uw_total += uw
            if u.contract_rent:
                ltl_total += uw - u.contract_rent

    return Totals(
        occ_units=occ_units,
        tot_units=tot_units,
        sqft=sqft,
        avg_sf=sqft / tot_units if tot_units else 0.0,
        mkt_rent=mkt_rent,
        avg_mkt_unit=mkt_rent / tot_units if tot_units else 0.0,
        avg_mkt_sf=mkt_rent / sqft if sqft else 0.0,
        cont_rent=cont_rent,
        avg_cont_unit=cont_rent / occ_units if occ_units else None,
        avg_cont_sf=cont_rent / occ_sqft if occ_sqft else None,
        max_rent=max_rent,
        other_inc=other_inc,
        occ_sqft=occ_sqft,
        occ_market_rent=occ_market,
        in_place_rent=in_place,
        pct_of_market=(in_place / occ_market) if (in_place and occ_market) else None,
        window_counts=window_counts,
        conc=conc,
        emp_disc=emp_disc,
        uw_market_total=uw_total,
        ltl_total=ltl_total,
        ltl_pct_total=(ltl_total / uw_total) if uw_total else None,
    )


def loss_to_lease(unit: Unit, config: RollConfig):
    """Return (uw_market, ltl, ltl_pct) for a Pres. RR row.

    LTL is only meaningful for units carrying a contract rent; vacant units
    (contract 0) show 0 / 0% to match the reference template.
    """
    uw = config.uw_market(unit.floor_plan)
    if uw is None:
        return None, None, None
    if unit.contract_rent:
        ltl = uw - unit.contract_rent
        return uw, ltl, (ltl / uw if uw else 0.0)
    return uw, 0, 0
