# Iteration Log

A running record of the deals processed and the tuning decisions made, so
future model tuning has the full context. Each decision notes *what* changed,
*why*, and *where* it lives in the code. Commit hashes are in `git log`.

---

## Deals processed

| # | Deal | Source format | As-of | Result |
|---|------|---------------|-------|--------|
| — | Heritage Hill Estates | OneSite-style (reference exhibits) | 03/16/26 | Reverse-engineered → regression fixture; engine reproduces it value-for-value |
| 1 | Station J Town | Yardi (`Report1` sheet) | 05/22/26 | 384 units, 94.27% occ; ties to source summary (Mkt 430,031 / Actual 385,359 / SF 365,100) |

---

## Tuning decisions (chronological)

1. **Output = 4 summary tabs, values-only** — matched the Heritage reference
   exactly (Unit Mix & RR Summary, Recent Leases, Unit Mix by Beds,
   Pres. Rent Roll). No live formulas in these tabs. Validated cell-for-cell
   via `tests/validate_heritage.py` (`ALL VALUES MATCH`).

2. **Business rules locked from the reference** (`aggregate.py`):
   - Occupancy: `Occ` occupied; `NTV`/`VL` → occupied; `Vacant` vacant.
   - Max Rent (Unit Mix) = highest **contract** rent among occupied units.
   - Recent-lease windows = trailing **calendar months** (180→6, 120→4, 90→3,
     60→2, 30→1), not raw day counts.
   - Loss-to-Lease = UW market − contract; 0 for vacant.

3. **Model occupancy made configurable** — `RollConfig.model_occupied`,
   **default False** (a model is non-revenue). Heritage opts in (`True`) to keep
   matching its reference (which folds the model into occupied). Station J Town
   uses the default → 94.27%, matching the source. *(schema.py, aggregate.py)*

4. **Lease-start fallback = Move-In date** — when a source has no lease-sign
   date (Station J Town has only Move-In + Lease Expiration), use Move-In as the
   lease-start placeholder for the Recent Leases windows. *(intake_station_jtown.py)*

5. **Other income = 0 when the source has no such column** (Station J Town).

6. **LTL left blank** — `uw_market_rents={}` per request; the AE/LTL/LTL% columns
   render empty. Fill `RollConfig.uw_market_rents` per floor plan to populate.

7. **OneLineRR tab added** (5th tab) — reproduces the analyst's working sheet:
   a full one-line-per-unit dump of the source plus a yellow, user-editable
   **Checking** table (unit type → floor plan / BD / BA / Renovated). Derived
   columns and COUNTIFS/SUMIFS aggregates are **live formulas** with
   dynamically sized ranges. *(workbook.py `_build_onelinerr`)*

8. **Station J Town unit-type decode** (from the analyst's Checking mapping):
   letter A=1BD, B=2BD, C=3BD; trailing digit 1=1BA, 2=1.5BA; `R` suffix =
   renovated. Applied to **all** tabs. *(intake_station_jtown.py `STJT_MAP`)*

9. **Output filename = modification date** — `Rent_Roll_Exhibits_<Deal>_mm.dd.yy`
   where `mm.dd.yy` is today's date, not the rent-roll date. *(intake_station_jtown.py)*

---

## Open / future tuning items

- **Model occupancy consistency** — the OneLineRR `# Occ` formula counts the
  model as occupied (reproduces the template), while the summary tabs exclude it
  (per decision #3). Currently intentional; align if desired.
- **`stjt2C2` bath** — set to 1.5 for consistency with its "3 BD / 1.5 BA"
  label; the analyst's sheet had 1.0 there. Yellow cell is user-editable.
- **Floor-plan / BD-BA decode is deal-specific** — lives in the per-source
  intake (`STJT_MAP`). Each new PM format needs its own mapping (or the user
  edits the OneLineRR Checking table).
- **UW market rents** — supply per floor plan to turn on the LTL columns.
- **Reno grouping** — summary tabs treat reno as a per-unit flag (grouped with
  classics); OneLineRR lists reno as a distinct `..., R` line. Revisit if a
  reno-split Unit Mix is wanted.

---

## Adding a new deal

1. Write a small intake mapper (copy `examples/intake_station_jtown.py`): read
   the source, map columns to the normalized `Unit` schema (see `PROCESS.md §4`),
   and provide a `RollConfig`.
2. Run it → 5-tab exhibits workbook.
3. Cross-check the totals against the source's own summary rows.
4. Record the deal and any new decisions in this log.
