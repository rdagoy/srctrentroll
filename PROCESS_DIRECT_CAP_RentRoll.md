# DIRECT CAP — Compact Rent Roll (raw RR → OneLineRR)

Process for mapping any property-management rent-roll export into the **OneLineRR**
tab of the DIRECT CAP UW Model. This is the SRCT "Multifamily Rent Roll (RR → UW
Model)" process, calibrated to the DIRECT CAP template (which has a different layout
than SRCT's base `UW_Template_base.xlsx`: data rows start at row 99, property name in
`$C$67`, and a built-in reconciliation to the raw-source footer).

> Source files are provided **separately per deal** — the workbook is a template only.

## Assets in this repo
| File | Purpose |
|---|---|
| `templates/DIRECT_CAP_UW_Model_TEMPLATE.xlsx` | Blank shell. OneLineRR inputs, floor-plan table, Property-1 header, and the `HVP RR` source tab are cleared. All formulas, 8 charts, 5 drawings, and data validations preserved. |
| `tools/populate_onelinerr.py` | Reusable engine: parses a source RR and writes OneLineRR into a copy of the template. |
| `PROCESS_DIRECT_CAP_RentRoll.md` | This document. |

## Non-negotiable rules (SRCT)
- **Never invent data.** Every written value comes from the source file.
- The raw unit-type string is preserved **exactly** in OneLineRR col K (the join key).
- Control totals are **reconciled against the source** before delivery.
- Imputed/blank fields are called out, never silent.
- No em dashes; institutional third person.

## OneLineRR template layout (DIRECT CAP)
- **Headers:** row 98. **Data rows:** 99–606 (508-unit scaffold; extend if a deal is larger).
- **Floor-plan lookup table:** rows 5–61. Per distinct unit type: `C` label, `D` raw unit
  type (join key), `E`/`F` BD/BA, `G` Affordable, `H` Renovated, `I` Reno Type. Columns
  `J–W` are COUNTIFS/SUMIFS rollups (do not edit). The derived columns `D–I` on each data
  row INDEX/MATCH into this table on (unit type × property).
- **Property header:** `C67` name, `C68` address, `C69` city/state/zip, `C70` as-of date.
- **Summary property link:** the summary property-name column `B5:B61` is set to `=$C$67`,
  so it always reflects the header name entered in `C67` (single-property deals).
- **Reconciliation:** `K62='HVP RR'!B526` compares the SUMMARY unit count against the raw
  source footer. Re-point this to the new source when the source tab is renamed.

## Column mapping (source → OneLineRR)
| OneLineRR | Col | Source field (default) | Notes |
|---|---|---|---|
| Unit # | J | Unit | text, verbatim |
| Unit Type | K | Unit Type | verbatim; must match a floor-plan table row |
| Unit Sq Ft | L | Sqft | number |
| Resident/Tenant | M | Tenant | vacant units read "Vacant" → drives Occ/Vac formula |
| Market Rent | N | market col, else placeholder | if the source has no market rent, N is imputed = max in-place rent of the same floor plan (see Imputation rules) |
| Rent | AA | Rent | feeds `O` Contract Rent `=SUM(AA:AB)` |
| Other Income | AD | Monthly Charges | feeds `T` Other Income `=SUM(AD:AJ)` |
| Lease Start | Q | Lease From | date |
| Lease End | R | Lease To | date |
| Move in / Move out | P / S | *(blank)* | not in typical export |
| Subsidy | AB | *(blank)* | not in typical export |

Derived by template formula (do not write): `A` Property, `B` counter, `C` Occupancy,
`D–I` Floorplan/Bd/Ba/Affordable/Renovated/Reno Type, `O/T/U/V/W` roll-up sums.

## Imputation rules (all deals)
- **Placeholder market rent.** When a unit has no market rent in the source, col N is
  filled with the **max in-place rent (Contract Rent = Rent + Subsidy) of the same floor
  plan**, applied to every unit of that floor plan. Units whose floor plan has no positive
  in-place rent (e.g. an on-site Office at $0) are left blank. Deals that supply a real
  market rent keep it; only blanks are imputed. These are placeholders — the engine reports
  the count on every run. Disable per deal with `MARKET_PLACEHOLDER=False`.

## Standing decisions (Hudson View Park calibration)
1. **Market Rent (col N).** The HVP export carries two figures — `Market Rent` and
   `Computed Market Rent` (amenity-loaded); pick one to map before any loss-to-lease work.
   Until one is mapped, the all-deals placeholder rule fills N with the max in-place rent
   of each floor plan. *(Column choice pending client instruction.)*
2. **Rent → col AA** (Contract Rent = Rent + Subsidy).
3. **Monthly Charges → col AD** (Other Income).
4. **Subsidy / Move-in / Move-out: blank** (absent from the export).
5. **Floor-plan table: rebuilt per deal** — human gate to confirm BD/BA and the
   Renovated flag (auto-default is `No`/`Classic`; the source's amenity/renovation
   signals are **not** auto-classified).

## Run sequence (every deal)
1. **Receive the source file** (provided separately, per deal).
2. Set the `CONFIG` block in `populate_onelinerr.py`:
   - `COLMAP` — match your export's header text (case-insensitive). Set a field to `None`
     to leave that OneLineRR column blank.
   - `PROPERTY` / `ADDRESS` / `CITY` / `ASOF` — the header block written to **C67 / C68 /
     C69 / C70**.
   - `FLOOR_PLANS` — leave `None` to auto-derive from the source (BD/BA parsed from
     `BD_BA_COL`); or supply an explicit list to set labels, BD/BA, and Renovated by hand.
3. **Run the engine:**
   ```
   python tools/populate_onelinerr.py  <SOURCE_RR.xlsx>  --out "<Deal> - UW Model.xlsx"
   ```
   It writes the data rows, the floor-plan table, the property block (C67:C70), and links
   `B5:B61` to `=$C$67`. Only `OneLineRR` and `workbook.xml` are touched (surgical XML
   edit), so charts, drawings, and validations are preserved; `fullCalcOnLoad` makes every
   formula recompute on open.
4. **CONFIRMATION GATE — clarify with the client first.** Before the file is considered
   final, present the reconciliation table the engine prints (source vs. workbook: units,
   vacant, rent total, other-income total, sqft total) and **confirm with the client that
   every gathered value matches the source.** Do not hand off until this is confirmed.
   Any `*** MISMATCH ***` must be resolved first.

## Verification (must pass before hand-off)
The engine reconciles the written-back workbook against the parsed source and prints a
`source` vs. `workbook` table with an `OK` / `MISMATCH` flag per field (`units`, `vacant`,
`rent_total`, `other_total`, `sqft_total`) plus the floor-plan count; it exits non-zero on
any mismatch. Also cross-check against the source's own footer totals when present.
In the workbook: OneLineRR Check row all TRUE / Diff row all 0. Spot-check vacant and any
placeholder rows. Confirm every unit type resolved (no `#N/A` in `D–I`), which means every
source unit type has a matching floor-plan table row.

## Source formats
- **Flat** (one row per unit): the default `CONFIG` / `parse_source` path.
- **Itemized** (one unit block spanning several charge rows ending in `Net:`): set
  `FORMAT="itemized"` and an `ITEMIZED` block (charge-column letters, header text, and the
  Rent / Subsidy / Commercial charge-name buckets). Charges bucket to Rent (AA), Subsidy
  (AB → e.g. Section 8), and Other Income (AD = everything else residential). Commercial
  units (by charge name) are excluded; `EXCLUDE_NONREV=True` also drops $0 non-revenue
  units (default keeps them). See `tools/deals/tradewinds.py` for a worked example.

Per-deal configs live in `tools/deals/<name>.py` and are passed with `--config`. When the
source has no BD/BA, the floor-plan table is written with labels sorted by average unit
sqft and BD/BA left blank for the analyst/client to fill (a hard human gate — never
inferred silently).

## Note on scope
Only the rent-roll surface is templated here (OneLineRR + its `HVP RR` source tab).
The other UW exhibits (P&Ls, Unit Mix, Comps, Tax, T12) are fed by other SRCT
processes and were left intact in the template.

## Improving this process
Per SRCT policy, submit process changes on the **SRCT Process Hub** (the "Improve"
box) so they apply firm-wide and are credited to you — the synced
`srct-rent-roll-multifamily` skill re-syncs from there.
