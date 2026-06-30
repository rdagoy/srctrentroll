"""Normalized input schema for the rent roll processor.

The processor is split in two halves on purpose:

  1. Intake / normalization (adaptive, per source format) produces a list of
     ``Unit`` objects + a ``RollConfig``.  This is the part that changes per
     PM system (OneSite, Yardi, RealPage, AppFolio, plain CSV, ...).
  2. The deterministic engine (``aggregate`` + ``workbook``) consumes that
     normalized data and always emits the same exhibits layout.

Keeping a single, documented normalized schema means the engine never has to
know which PM system a deal came from.
"""
from __future__ import annotations

from dataclasses import dataclass, field, fields
from datetime import date, datetime
from typing import Optional, Dict, Any


# --- Occupancy normalization -------------------------------------------------
# Raw status strings from PM exports get mapped to one of three buckets.
# Per underwriting convention (SOP v3): NTV (notice-to-vacate) and VL
# (vacant-leased) both count as OCCUPIED.  Model units are non-revenue and are
# NOT occupied by default (configurable via RollConfig.model_occupied).
OCC = "Occ"
VAC = "Vac"
MODEL = "Model"

# Default raw -> bucket mapping (compared case-insensitively, stripped).
DEFAULT_STATUS_MAP: Dict[str, str] = {
    "occ": OCC,
    "occupied": OCC,
    "current": OCC,
    "ntv": OCC,            # notice to vacate -> still occupied
    "ntvl": OCC,
    "notice": OCC,
    "notice-rented": OCC,
    "notice unrented": OCC,
    "vl": OCC,             # vacant-leased -> counts as occupied
    "vacant-leased": OCC,
    "vacant leased": OCC,
    "vacant rented": OCC,
    "vac": VAC,
    "vacant": VAC,
    "vacant-unrented": VAC,
    "vacant unrented": VAC,
    "vr": VAC,
    "down": VAC,
    "model": MODEL,
    "mdl": MODEL,
}


def normalize_status(raw: Any, status_map: Optional[Dict[str, str]] = None) -> str:
    """Map a raw occupancy string to Occ / Vac / Model."""
    smap = status_map or DEFAULT_STATUS_MAP
    if raw is None:
        return VAC
    key = str(raw).strip().lower()
    if key in smap:
        return smap[key]
    # Unknown -> assume occupied if it looks like a person/tenant tag, else vac.
    return smap.get(key, OCC if key else VAC)


def _coerce_date(value: Any) -> Optional[date]:
    if value in (None, "", "n/a", "N/A"):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    s = str(value).strip()
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%d-%b-%Y", "%d-%b-%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _num(value: Any, default: float = 0.0) -> float:
    if value in (None, ""):
        return default
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip().replace("$", "").replace(",", "")
    if s in ("", "-", "n/a", "N/A"):
        return default
    paren = s.startswith("(") and s.endswith(")")
    if paren:
        s = s[1:-1]
    try:
        v = float(s)
    except ValueError:
        return default
    return -v if paren else v


@dataclass
class Unit:
    """One residential unit, normalized.

    Column letters in comments map to the Pres. Rent Roll tab so it is easy to
    cross-check against the exhibits template.
    """
    floor_plan: str                       # G  - grouping key, e.g. "2 BD / 2 BA - A"
    sqft: float                           # L  - net square feet
    beds: float                           # M
    baths: float                          # N
    market_rent: float                    # U  - per-unit market/asking rent
    contract_rent: float = 0.0            # V  - in-place rent (0 if vacant)

    unit_id: str = ""                     # E  - unit number
    unit_type: Optional[str] = None       # F  - PM unit-type code (e.g. "220")
    property_name: Optional[str] = None   # D  - usually filled from config
    designation: Optional[str] = None     # H
    renovated: Optional[str] = None       # I  - "Yes" / None
    reno_type: Optional[str] = None       # J  - e.g. "Full"
    tenant_name: Optional[str] = None     # K  - "VACANT"/"MODEL" allowed
    lease_type: Optional[str] = None      # O
    reno_status: Optional[str] = None     # P
    occupancy: str = OCC                  # Q  - normalized Occ/Vac/Model
    mtm: Optional[str] = None             # R
    renew_status: Optional[str] = None    # S
    asking_rent: Optional[float] = None   # T
    concession: float = 0.0               # W  - conc/special (negative)
    emp_discount: float = 0.0             # X
    other_income: float = 0.0             # Y
    move_in: Optional[date] = None        # Z
    lease_start: Optional[date] = None    # AA
    lease_end: Optional[date] = None      # AB

    # --- Derived helpers -----------------------------------------------------
    @property
    def is_occupied(self) -> bool:
        """Simple per-unit helper (Occ only). Whether *model* units count as
        occupied is decided at aggregation time by RollConfig.model_occupied,
        so summaries use aggregate.is_occupied(unit, model_occupied) instead."""
        return self.occupancy == OCC

    @property
    def is_vacant(self) -> bool:
        """True for vacant units (used only for the Pres. RR tenant label).
        Model units are not 'vacant' here — they carry their own MODEL label."""
        return self.occupancy == VAC

    @classmethod
    def from_dict(cls, d: Dict[str, Any], status_map: Optional[Dict[str, str]] = None) -> "Unit":
        """Build a Unit from a loose dict (e.g. parsed JSON / a source row).

        Numbers and dates are coerced; occupancy is normalized.  Unknown keys
        are ignored so callers can pass through extra source columns.
        """
        valid = {f.name for f in fields(cls)}
        data = {k: v for k, v in d.items() if k in valid}

        for k in ("sqft", "beds", "baths", "market_rent", "contract_rent",
                  "concession", "emp_discount", "other_income"):
            if k in data:
                data[k] = _num(data[k])
        if "asking_rent" in data and data["asking_rent"] not in (None, ""):
            data["asking_rent"] = _num(data["asking_rent"])
        for k in ("move_in", "lease_start", "lease_end"):
            if k in data:
                data[k] = _coerce_date(data[k])

        raw_occ = d.get("occupancy", d.get("occ"))
        data["occupancy"] = normalize_status(raw_occ, status_map)

        # Vacant units carry no contract rent.
        if data.get("occupancy") == VAC:
            data["contract_rent"] = 0.0
        return cls(**data)


@dataclass
class RollConfig:
    """Per-deal configuration that the engine needs.

    Attributes:
        property_name: Deal name, shown in the title block of every tab.
        as_of_date:    Rent roll "as-of" date.  Drives the Recent Leases
                       rolling windows and the RR-date box.
        uw_market_rents: floor_plan -> underwriting market rent.  Populates the
                       AE "Market Rent" column on Pres. Rent Roll and the LTL
                       calc (LTL = uw_market - contract_rent).  A plan missing
                       here gets a blank AE / zero LTL.
        reno_tiers:    optional reno-code -> display-name map (documentation /
                       future use; the current exhibits format does not color
                       reno rows).
        status_map:    optional override of the raw->Occ/Vac/Model mapping.
        bed_total_max_rent: what to put in the bed-mix Max Rent total cell.
                       The reference file uses the literal string "N/A".
        model_occupied: whether model units count as occupied.  Defaults to
                       False (the standard underwriting view: a model is a
                       non-revenue unit).  Set True to match exhibits that
                       fold the model into the occupied count.
    """
    property_name: str
    as_of_date: date
    uw_market_rents: Dict[str, float] = field(default_factory=dict)
    reno_tiers: Dict[str, str] = field(default_factory=dict)
    status_map: Optional[Dict[str, str]] = None
    bed_total_max_rent: Any = "N/A"
    model_occupied: bool = False

    def __post_init__(self):
        self.as_of_date = _coerce_date(self.as_of_date) or self.as_of_date

    def uw_market(self, floor_plan: str) -> Optional[float]:
        return self.uw_market_rents.get(floor_plan)
