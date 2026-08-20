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
- **Reconciliation:** `K62='HVP RR'!B526` compares the SUMMARY unit count against the raw
  source footer. Re-point this to the new source when the source tab is renamed.

## Column mapping (source → OneLineRR)
| OneLineRR | Col | Source field (default) | Notes |
|---|---|---|---|
| Unit # | J | Unit | text, verbatim |
| Unit Type | K | Unit Type | verbatim; must match a floor-plan table row |
| Unit Sq Ft | L | Sqft | number |
| Resident/Tenant | M | Tenant | vacant units read "Vacant" → drives Occ/Vac formula |
| Market Rent | N | *(blank)* | **left blank by default** — see decision below |
| Rent | AA | Rent | feeds `O` Contract Rent `=SUM(AA:AB)` |
| Other Income | AD | Monthly Charges | feeds `T` Other Income `=SUM(AD:AJ)` |
| Lease Start | Q | Lease From | date |
| Lease End | R | Lease To | date |
| Move in / Move out | P / S | *(blank)* | not in typical export |
| Subsidy | AB | *(blank)* | not in typical export |

Derived by template formula (do not write): `A` Property, `B` counter, `C` Occupancy,
`D–I` Floorplan/Bd/Ba/Affordable/Renovated/Reno Type, `O/T/U/V/W` roll-up sums.

## Standing decisions (Hudson View Park calibration)
1. **Market Rent (col N): left blank.** The export carries two disagreeing figures —
   `Market Rent` and `Computed Market Rent` (amenity-loaded). Choose one before any
   loss-to-lease work. *(Pending client instruction.)*
2. **Rent → col AA** (Contract Rent = Rent + Subsidy).
3. **Monthly Charges → col AD** (Other Income).
4. **Subsidy / Move-in / Move-out: blank** (absent from the export).
5. **Floor-plan table: rebuilt per deal** — human gate to confirm BD/BA and the
   Renovated flag (auto-default is `No`/`Classic`; the source's amenity/renovation
   signals are **not** auto-classified).

## How to run
```
python tools/populate_onelinerr.py  <SOURCE_RR.xlsx>  --out "<Deal> - UW Model.xlsx"
```
Per deal, edit the `CONFIG` block at the top of `populate_onelinerr.py`:
- `COLMAP` — match your export's header text (case-insensitive). Set a field to `None`
  to leave that OneLineRR column blank.
- `PROPERTY` / `ADDRESS` / `CITY` / `ASOF` — header block.
- `FLOOR_PLANS` — leave `None` to auto-derive from the source (BD/BA parsed from
  `BD_BA_COL`); or supply an explicit list to set labels, BD/BA, and Renovated by hand.

The engine writes only `OneLineRR` and `workbook.xml` (surgical XML edit), so charts,
drawings, and validations are preserved. `fullCalcOnLoad` is set so every formula
recomputes when Excel opens the file.

## Verification (must pass before hand-off)
The engine prints a reconciliation report — check it against the source footer:
`units`, `floor_plans`, `vacant`, `rent_total`, `other_total`, `sqft_total`.
In the workbook: OneLineRR Check row all TRUE / Diff row all 0. Spot-check vacant and
any placeholder rows. Confirm every unit type resolved (no `#N/A` in `D–I`), which
means every source unit type has a matching floor-plan table row.

## Note on scope
Only the rent-roll surface is templated here (OneLineRR + its `HVP RR` source tab).
The other UW exhibits (P&Ls, Unit Mix, Comps, Tax, T12) are fed by other SRCT
processes and were left intact in the template.

## Improving this process
Per SRCT policy, submit process changes on the **SRCT Process Hub** (the "Improve"
box) so they apply firm-wide and are credited to you — the synced
`srct-rent-roll-multifamily` skill re-syncs from there.
