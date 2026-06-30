# Rent Roll Processor

Turn a raw property-management rent roll (OneSite, Yardi, RealPage, AppFolio,
plain Excel/CSV, …) into the standardized **Rent Roll Exhibits** workbook used
for underwriting.

The output is a 4-tab `.xlsx` that matches the reference exhibits layout
cell-for-cell:

| Tab | One row per | Contents |
|-----|-------------|----------|
| **Unit Mix & RR Summary** | floor plan | occ/total units, avg rents, $/SF, max rent, other income, + a summary-stat block and an audit table |
| **Recent Leases** | floor plan | trailing 6/4/3/2/1-month average lease rates and lease counts |
| **Unit Mix based on # of Beds** | bed count | same columns as Unit Mix, grouped by bedroom count |
| **Pres. Rent Roll** | unit | every unit with status, rents, dates, and loss-to-lease (LTL) |
| **OneLineRR** | unit | full one-line dump of source data + a user-editable "Checking" table (unit-type → floor-plan / BD / BA / Renovated mapping) with live COUNTIFS/SUMIFS aggregates |

## How the process works

The processor is deliberately split into two halves (see
[`PROCESS.md`](PROCESS.md) for the full SOP):

1. **Intake / normalization** *(adaptive — changes per source format)*
   The raw source is read and each unit is mapped to the normalized
   [`Unit`](rent_roll_processor/schema.py) schema. This is the part that knows
   about OneSite vs. Yardi column names, reno-code suffixes, etc.
2. **Exhibits engine** *(deterministic — always the same output)*
   The normalized units + a [`RollConfig`](rent_roll_processor/schema.py) are
   fed to `build_exhibits()`, which computes every summary and writes the
   styled workbook.

Keeping a single normalized schema means the engine never has to care which PM
system a deal came from — only the intake mapping changes.

## Usage

```bash
pip install -r requirements.txt

# From a single deal file (config + normalized units):
python -m rent_roll_processor.cli --deal examples/sample_deal.json -o exhibits.xlsx

# Or as a library:
python - <<'PY'
from datetime import date
from rent_roll_processor import Unit, RollConfig, build_exhibits

config = RollConfig(
    property_name="Sample Apartments",
    as_of_date=date(2026, 1, 31),
    uw_market_rents={"1 BD / 1 BA": 1250, "2 BD / 2 BA": 1500},
)
units = [Unit.from_dict(row) for row in my_rows]   # my_rows = normalized dicts
build_exhibits(units, config, "exhibits.xlsx")
PY
```

## Validation

The engine is regression-tested against the original Heritage Hill Estates
exhibits: the 120 units are re-fed through the engine and **every value** in
all four tabs is diffed against the reference workbook.

```bash
python -m tests.validate_heritage /path/to/Rent_Roll_Exhibits_Heritage_Hill...xlsm
# -> ALL VALUES MATCH
```

## Key business rules (reverse-engineered from the reference)

- **Occupancy:** `Occ` counts as occupied; `NTV` (notice-to-vacate) and `VL`
  (vacant-leased) also normalize to occupied. `Vacant` is vacant. **Model**
  units are *not* occupied by default (`RollConfig.model_occupied=True` folds
  them in, as the Heritage reference does).
- **Max Rent** (Unit Mix) = the highest *in-place* (contract) rent among
  occupied units in the group.
- **Recent-lease windows** are *calendar months* (1/2/3/4/6), not raw day
  counts — a lease 61 days out still lands in the "60 Days" column.
- **Loss to Lease** = `underwriting market rent − contract rent`; vacant units
  (no contract) show `0`. The underwriting market rent comes from
  `RollConfig.uw_market_rents` (an analyst input per floor plan).
- All money/summary cells are written as **values** (the reference exhibits
  contain no live formulas).

See [`PROCESS.md`](PROCESS.md) for the normalized schema reference, the QC
checklist, and how to run a new deal.
