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
| 2 | Maven @ 806 | 29SC / RealPage (`Detailed` charge-ledger) | 05/15/26 | 51 units, 82.35% occ; ties to source totals (SF 38,376 / scheduled charges 58,021.30 = contract 52,035.45 + other 6,361 + conc −374.70) |
| 3 | Colonial Pointe (556 + 558) | **scanned/image PDF** (Q1 2026, pages 2 & 4) | 03/31/26 | 88 units (65 + 23), 2 vacant; rent ties: 556 = 218,106 exact, 558 line items = 33,498 (source printed 33,500 — a $2 source artifact) |
| 4 | Portage Towers | Berkadia manual Excel ('April' book) | 04/30/26 | 378 units (2 towers), 96.03% occ; ties **exactly** to source grand totals (SF 332,650 · Market 421,990 · Base 375,130 · Discount −20,680 · Other 35,241.95) |
| 5 | Craigdell Gardens | AppFolio Rent Roll export | 07/13/26 | 97 units, 91.75% occ; ties **exactly** to source totals (Units 97 · SF 76,200 · Market 85,460 · Rent 79,227 · Monthly Charges 1,590) |

### Lessons (Colonial Pointe)
* **Image-only PDF** → no extractable text; render pages to PNG (PyMuPDF) and
  read them, then transcribe. Rotate to upright first.
* **Two buildings, one property** → combined into one deal; unit ids prefixed
  with building (`556-201`, `558-101`) to stay unique.
* **Unknown BD/BA** (source has none) → `beds/baths=None`, floor plan = the
  type code, sorted by SF (rule #21 fallback); fill the OneLineRR Checking
  table to regroup. Fixed `Unit.from_dict` so `None` beds/baths stay `None`
  (were coerced to 0) and `bed_mix` labels a `None` group "N/A".
* **Single rent figure** ("Total") → used as both market and in-place rent;
  no other income / concession; **no lease-start/move-in** so Recent Leases is
  empty. Note-flagged units (Leasing Mgr Storage, Maintenance Super Apt) are
  not auto-detected — surface them for the analyst to classify.

### Lesson (Maven): unit-count sanity check
Maven's Bldg-Unit ids came in **two formats** — `800-1A` and `CL - 806-1`. The
first intake keyed unit rows off the id regex `^\d+-`, silently skipping the 37
`CL - ` units → 14 instead of 51. Fix: detect unit header rows by the **Unit
Type** column (`\dx\d`), not the id format. **Always cross-check the output unit
count against the source's own total** (e.g. "Total Rentable Units").

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

9. **Output filename = modification date** — `mm.dd.yy` is today's date, not the
   rent-roll date. *(superseded by #16)*

10. **Vacant market rent derived** — each vacant unit's market rent is set to the
    in-place (contract) rent of the most-recently-started lease of the same
    **unit type** (strictly the unit-type code — see #13/#14); fallback is that
    unit type's max in-place rent; if no occupied comp exists the source value is
    kept. Toggle `RollConfig.derive_vacant_market` (default True).
    *(derive.py `_apply_vacant_market`)*

11. **Non-revenue units expanded** — a unit whose tenant name is
    `admin`/`down`/`super`/`model` (or whose status already maps to one) is
    normalized to that label, its rent is set = market rent, and an offsetting
    negative concession is added (net rent 0). Non-revenue units are not counted
    as occupied. Toggle `RollConfig.expand_nonrev` (default True).
    *(derive.py `_apply_nonrev`; new statuses Admin/Down/Super in schema.py;
    OneLineRR occupancy formula + concession column updated)*
    - *Station J Town effect:* model E06 → rent 1135 (= market), concession -1135.

    Both derivations run in `build_exhibits` (via `derive.apply_deal_rules`)
    before aggregation, so they flow to every tab. The Heritage fixture sets both
    toggles False to stay a frozen reference.

12. **In-place rent netted by concession** — the summary tabs' in-place/contract
    rent (Unit Mix Cont Rent + averages, Recent Leases in-place + windows,
    In-Place Annualized, totals) now use `contract + concession` so a
    non-revenue unit nets to 0 and real concessions reduce in-place rent. The
    Pres. Rent Roll still shows gross contract and concession separately. Toggle
    `RollConfig.net_concession` (default True; Heritage sets False).
    *(aggregate.py `in_place_rent`)*
    - *Station J Town effect:* 2 BD / 1 BA Cont Rent 207,560 (gross) → 206,425
      (net); the model's 1,135 nets out.

13. **Vacant-market fix: reno-aware comps** — the vacant-market calc (#10)
    originally grouped comps by merged floor plan, so a *classic* vacant could
    pick up a *renovated* unit's rent (and vice-versa). First fixed by bucketing
    on `(floor_plan, renovated)`.

14. **Vacant-market: strict unit-type basis** — per clarification, the comp
    bucket is now **strictly the unit-type code** (`_comp_key` returns
    `unit.unit_type`), not floor plan or renovation status. *(derive.py `_comp_key`)*
    - *Verified:* all 21 Station J Town vacants match the strict unit-type rule
      (e.g. stjt2B1 → 965, stjt2B2R → 1190). Requires each Unit to carry its raw
      `unit_type` code.

15. **Occupied $0-rent placeholder** — an occupied unit carrying $0 contract rent
    uses its market rent as a placeholder in-place rent (no concession). Runs
    after the vacant-market calc so the $0 unit is never used as a vacant comp;
    vacant and non-revenue units are untouched. Toggle
    `RollConfig.zero_rent_placeholder` (default True; Heritage sets False).
    *(derive.py `_apply_zero_rent_placeholder`)*
    - *Station J Town effect:* L30 (Jose Mendoza Hernandez, stjt2A1) contract
      0 → 990 (= market). All 21 vacants still verify.

16. **Output filename convention** — every deal is named
    `Birgo RR Exhibits - <Property Name>v<mm.dd.yy>.xlsx`, where `v` = "version"
    followed by today's (file-modified) date. *(naming.py `output_filename`;
    used by the CLI when `--out` is omitted and by the intake scripts)*
    - e.g. `Birgo RR Exhibits - Station J Townv07.01.26.xlsx`

18. **OneLineRR: itemized Other Income** — the Other Income section (columns
    right of Rent, from AB) now renders **one column per other-income line item**
    with per-unit values; column R = SUM over those columns (Checking "Other
    Income" aggregates it). Falls back to a single "Other Income" column when a
    source has no per-item breakdown. Concessions get their own column just after.
    *(schema `Unit.other_income_items`; intake populates it; workbook
    `_build_onelinerr` places columns dynamically)*
    - *Maven effect:* 8 columns (Building Protection, Pet, Pest, Trash, Water,
      WiFi, MTM Fee, RUBS); detail totals 6,361 = source other income.

24. **Other Income: pet charges leftmost** — in the OneLineRR Other Income
    section, any line item whose name contains "pet" is placed in the leftmost
    column(s); the rest keep their first-seen order (stable partition).
    *(workbook.py `_build_onelinerr`)*
    - *Maven:* "Pet Rent" now the first Other Income column (AB).

19. **Vacant name normalization** — any tenant name containing "vacant" (e.g.
    "-- Vacant --", "VACANT") collapses to the literal "Vacant". Toggle
    `normalize_vacant_name` (default True; Heritage False). *(derive.py
    `_apply_vacant_name`)* — also fixes the OneLineRR occupancy formula, which
    tests `K="Vacant"` exactly.

20. **Vacant row highlight** — on the **OneLineRR** tab every vacant unit's data
    row is filled yellow (RGB 255,255,0) with blue font (RGB 0,0,255).
    *(workbook.py `VACANT_FILL`/`VACANT_FONT`)*

23. **Employee-discount routing** — any charge whose name contains "employee
    discount" is classified to the **employee-discount** field (Pres. RR col X;
    OneLineRR gets an "Employee Discounts" detail column, T = SUM). Not netted
    into in-place rent (netting stays concession-only). *(intake classification;
    workbook OneLineRR emp column)*
    - *Maven:* "Concession-Employee Rent Discount Special" (−374.70) moved from
      concession → employee discount on 800-2C. Grand scheduled total unchanged
      (58,021.30); that unit's netted in-place rent rises by 374.70.
    - **Decision (confirmed):** employee discounts are NOT netted into in-place
      rent — netting stays concession-only. Employee discount is treated as a
      separate, removable line (potential add-back), not a permanent rent cut.

22. **Move-out data on OneLineRR** — the `Unit` schema now carries `move_out`
    and the OneLineRR "Move Out" column (Q) is populated whenever the source has
    it. *(schema.py field + coercion; intakes map it — Maven "Expected Move-Out",
    Station J Town "Move Out"; workbook writes Q)*
    - *Maven:* 3 move-outs (the notice units). Always include when available.

21. **Unit-type ordering: beds → baths → SF** — the group sort (Unit Mix, Bed
    Mix, Recent Leases) and the OneLineRR Checking table now order by bed count,
    then bath count, then unit SF; unknown beds/baths fall back to SF (sort
    last). *(aggregate.py `group_sort_key`; workbook.py `_olr_unit_types`)*
    - Heritage order unchanged (SF is constant within each same-bed/bath plan),
      so the regression still matches.


25. **Lease-dates note** — when no unit has a lease-start date, the Recent Leases
    tab prints "Lease dates not available" in the cell directly below the
    Total/Wtd Average row (column C). *(workbook.py `_build_recent_leases`)*

26. **Admin (non-revenue) from notes** — units flagged by a note as storage /
    super / maintenance / office / leasing / employee / model are set to
    occupancy `Admin` (non-revenue) in the intake; the engine then applies
    rent = market + offsetting concession and excludes them from occupied.
    *(intake note keyword match; existing `expand_nonrev`)*
    - *Colonial Pointe:* 556-000 (Leasing Mgr Storage) and 558-307 (Maintenance
      Super Apt) -> Admin.

27. **Combined multi-building deal** — buildings 556 + 558 combined into one
    "Colonial Pointe" deal with building-prefixed unit ids (`556-201`).

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
