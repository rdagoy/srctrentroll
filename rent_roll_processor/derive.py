"""Deal-level derivations applied before aggregation.

These run inside ``build_exhibits`` (gated by RollConfig toggles) so every deal
gets the same treatment:

  * expand_nonrev        -> normalize non-revenue units and add the
                            rent = market / concession = -rent offset.
  * derive_vacant_market -> set vacant market rents from the latest-leased
                            in-place rent of the same floor plan.

Order matters: non-revenue units are normalized first (so they are excluded
from the occupied comps used for the vacant-market calc).
"""
from __future__ import annotations

import copy
from typing import List

from .schema import Unit, RollConfig, OCC, VAC, NON_REVENUE, normalize_status


def _nonrev_key(unit: Unit, status_map) -> str | None:
    """Return the non-revenue status for a unit, or None.

    A unit is non-revenue if its occupancy already maps to one of the
    non-revenue buckets, or if its tenant name is admin/down/super/model.
    """
    if unit.occupancy in NON_REVENUE:
        return unit.occupancy
    name = (unit.tenant_name or "").strip().lower()
    if name in ("admin", "down", "super", "model"):
        return normalize_status(name, status_map)
    return None


def _apply_nonrev(units: List[Unit], config: RollConfig) -> None:
    for u in units:
        key = _nonrev_key(u, config.status_map)
        if not key:
            continue
        u.occupancy = key
        u.tenant_name = key                 # canonical label (e.g. "Model")
        u.contract_rent = u.market_rent      # rent = market rent
        u.concession = -(u.market_rent or 0)  # offsetting negative concession


def _comp_key(unit: Unit):
    """Comp bucket for the vacant-market calc: the raw unit type, strictly.
    A vacant unit is priced only off occupied units of the exact same unit
    type (not merged floor plan, not renovation status)."""
    return unit.unit_type


def _apply_vacant_market(units: List[Unit]) -> None:
    # Build comps (floor plan + reno status) from truly occupied units only.
    from collections import defaultdict
    by_key = defaultdict(list)
    for u in units:
        if u.occupancy == OCC:
            by_key[_comp_key(u)].append(u)

    for u in units:
        if u.occupancy != VAC:
            continue
        comps = by_key.get(_comp_key(u), [])
        leased = [c for c in comps if c.lease_start is not None and c.contract_rent]
        if leased:
            # in-place rent of the most-recently-started lease (ties -> higher rent)
            latest = max(leased, key=lambda c: (c.lease_start, c.contract_rent))
            u.market_rent = latest.contract_rent
        else:
            # fallback: the bucket's max in-place rent among occupied units
            rents = [c.contract_rent for c in comps if c.contract_rent]
            if rents:
                u.market_rent = max(rents)
            # else: no occupied comp -> leave the source market rent as-is


def apply_deal_rules(units: List[Unit], config: RollConfig) -> List[Unit]:
    """Return a new list of units with the deal-level derivations applied.

    The input units are not mutated (each is shallow-copied)."""
    out = [copy.copy(u) for u in units]
    if config.expand_nonrev:
        _apply_nonrev(out, config)
    if config.derive_vacant_market:
        _apply_vacant_market(out)
    return out
