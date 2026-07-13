# Rent Roll Processing — Standard Operating Procedure

**Goal:** you send a raw source rent roll; you get back a standardized 5-tab
**Rent Roll Exhibits** workbook ready for underwriting.

This SOP describes how a new deal is run end to end, the data contract between
the two halves of the processor, and the QC checks to run before distributing.

---

## 1. What to send

For each deal, provide:

1. **The source rent roll** — any PM export (OneSite, Yardi, RealPage,
   AppFolio, Entrata, a plain Excel/CSV, …). Send it as-is; `.xls`, `.xlsx`,
   and `.csv` are all fine.
2. **Property name** — exactly as used in the underwriting model.
3. **As-of date** — the rent roll date from the source header (MM/DD/YYYY).
   *Use the date on the file, not today's date* — it drives the Recent Leases
   windows.
4. **Underwriting market rents per floor plan** — the market rent assumption
   used for the loss-to-lease (LTL) column. If you don't supply these, the LTL
   columns are left blank and can be filled in later. (You can also say
   "use the in-place high" or send a quick table.)
5. **Reno tiers**, if any — e.g. `VA1 = Renovated, VA2 = Premium Reno`, or
   "All Classic" if there's no reno program.

A copy/paste intake block:

```
PROPERTY:   <deal name>
AS-OF DATE: <MM/DD/YYYY>
UW MARKET RENTS:
  <floor plan> = <rent>
  ...
RENO TIERS: <VA1 = Renovated | ... | or "All Classic">
SOURCE FILE: attached
```

---

## 2. How a deal is processed

### Step 1 — Intake & normalization *(adaptive)*
The source columns are inspected and each unit is mapped to the normalized
`Unit` schema (Section 4). This step absorbs all format differences:

- Locate the header row and the unit rows (skip group/subtotal rows).
- Map source columns → normalized fields (unit #, floor plan, SF, beds/baths,
  market rent, in-place/contract rent, lease dates, status, other income).
- Normalize occupancy strings to **Occ / Vac / Model** (NTV and VL → Occ).
- Detect reno tiers from floor-plan suffixes if applicable.
- Coercions (numbers like `$1,234`, parenthesized negatives, multiple date
  formats) are handled automatically by `Unit.from_dict`.

The result is a list of normalized units + a `RollConfig`.

### Step 2 — Build exhibits *(deterministic)*
`build_exhibits(units, config, out_path)` computes every summary and writes the
styled 5-tab workbook (Unit Mix & RR Summary, Recent Leases, Unit Mix by
Beds, Pres. Rent Roll, and OneLineRR). Same input → identical output, every time.

The **OneLineRR** tab is a live working sheet: a full one-line-per-unit dump of
the source plus a yellow-highlighted **Checking** table that maps each unit
type to its floor plan / BD / BA / Renovated flag. Those mapping cells are
user-editable — edit them and the derived columns and aggregates recalculate.

```bash
python -m rent_roll_processor.cli --deal deal.json
# -> Birgo RR Exhibits - <Property Name>v<mm.dd.yy>.xlsx  (omit -o to use the convention)
```

### Step 3 — Confirmation summary
Along with the file you get: total units, occupied count + occupancy %, number
of unit types, and any data warnings (missing SF, stale/missing lease dates,
model units, large negative LTL, floor plans missing a UW market rent).

---

## 3. Business rules (baked into the engine)

| Rule | Behavior |
|------|----------|
| Occupancy | `Occ` counts as occupied; `NTV` and `VL` normalize to occupied. `Vacant` is vacant. **Model** units are **not** occupied by default (a non-revenue unit) — set `model_occupied=True` to fold them into occupied. |
| Occupied SF | sum of SF over occupied units. |
| Mkt Rent (group) | sum of per-unit market rent over **all** units. |
| Avg Mkt/Unit | mkt rent ÷ **total** units. Avg Mkt/SF = mkt rent ÷ total SF. |
| Vacant market rent | recomputed to the **in-place rent of the most-recently-started lease of the same *unit type*** (strictly the unit-type code — not floor plan, not renovation status); fallback: that unit type's max in-place rent. Toggle `derive_vacant_market`. |
| Non-revenue units | a unit named `admin`/`down`/`super`/`model` **or flagged offline by the source's own status/tag column** (e.g. an AppFolio `DOWN` tag) → status normalized to that label, **rent set = market rent**, and an offsetting **negative concession** added (net rent 0). Non-revenue units are excluded from **both** occupied and vacant. Toggle `expand_nonrev`. |
| Lease-dates note | if the source has no lease-start dates, the Recent Leases tab shows "Lease dates not available" in the cell just below the Total row. |
| Admin (non-revenue) units | units flagged (e.g. by a note) as storage / super / maintenance / office / leasing / employee / model are marked **Admin**: rent = market, offsetting negative concession, excluded from occupied. |
| Vacant name | any tenant name containing "vacant" (e.g. "-- Vacant --") is collapsed to the literal **"Vacant"**. Toggle `normalize_vacant_name`. |
| Vacant row highlight | on the **OneLineRR** tab, every vacant unit's row is filled yellow (255,255,0) with blue font (0,0,255). |
| Other Income order | in the OneLineRR Other Income section, **pet-related charges are placed leftmost** (any line item whose name contains "pet"); other items keep their order. |
| Unit-type ordering | summary rows and the OneLineRR Checking table sort by **bed count → bath count → unit SF** (SF is the basis when beds/baths are unknown). |
| Employee discount | any charge whose name contains "employee discount" is routed to the **Employee Discount** section (Pres. RR col X; OneLineRR "Employee Discounts"), not concession. It is *not* netted into in-place rent (netting is concession-only). |
| Occupied, $0 rent | an **occupied** unit carrying $0 contract rent uses its **market rent as a placeholder** in-place rent (no concession). Vacant/non-revenue units are unaffected. Toggle `zero_rent_placeholder`. |
| In-place / Cont Rent | **net of concession** in the summary tabs (contract + concession), so a non-revenue unit nets to 0 and real concessions reduce in-place rent. The Pres. Rent Roll still shows gross contract and concession in separate columns. Toggle `net_concession`. |
| Avg Cont/Unit | cont rent ÷ **occupied** units. Avg Cont/SF = cont rent ÷ **occupied** SF. |
| Max Rent | highest **contract** rent among occupied units in the group. |
| Recent-lease windows | trailing **calendar months** (180→6, 120→4, 90→3, 60→2, 30→1). Average of contract rent for occupied units whose lease start falls in the window; `n/a` when none. If the source has no lease-sign date, use the **Move-In date** as the lease-start placeholder. |
| Loss to Lease | `UW market rent − contract rent`; `0` for vacant (no contract). LTL % = LTL ÷ UW market. |
| Annualized | Market/In-Place/Other income × 12. |
| Output | the four summary tabs are values-only (match the reference exhibits); the OneLineRR tab is formula-driven so the Checking mapping recalculates live. |

---

## 4. Normalized `Unit` schema

One dict per unit. Only `floor_plan`, `sqft`, `beds`, `baths`, and
`market_rent` are strictly required; the rest default sensibly. Extra keys are
ignored. (Maps to the Pres. Rent Roll columns shown in parentheses.)

| Field | Type | Notes |
|-------|------|-------|
| `floor_plan` | str | **required** — grouping key, e.g. `"2 BD / 2 BA - A"` (col G) |
| `sqft` | number | **required** (col L) |
| `beds` / `baths` | number | **required** (cols M/N) |
| `market_rent` | number | **required** — per-unit market/asking rent (col U) |
| `contract_rent` | number | in-place rent; `0`/blank for vacant (col V) |
| `unit_id` | str | unit number (col E) |
| `unit_type` | str | PM unit-type code, e.g. `"220"` (col F) |
| `property_name` | str | usually filled from config (col D) |
| `designation` | str | (col H) |
| `renovated` | str | e.g. `"Yes"` (col I) |
| `reno_type` | str | e.g. `"Full"` (col J) |
| `tenant_name` | str | `"VACANT"`/`"MODEL"` allowed (col K) |
| `lease_type` | str | (col O) |
| `reno_status` | str | (col P) |
| `occupancy` | str | raw status; normalized to Occ/Vac/Model (col Q) |
| `mtm` | str | (col R) |
| `renew_status` | str | (col S) |
| `asking_rent` | number | (col T) |
| `concession` | number | conc/special, negative (col W) |
| `emp_discount` | number | (col X) |
| `other_income` | number | (col Y) |
| `move_in` | date | (col Z) |
| `lease_start` | date | drives Recent Leases windows (col AA) |
| `lease_end` | date | (col AB) |
| `move_out` | date | expected/actual move-out (OneLineRR Move Out) |

`RollConfig` fields: `property_name`, `as_of_date`, `uw_market_rents`
(`{floor_plan: rent}`), `reno_tiers`, `status_map` (optional override of the
raw→Occ/Vac/Model mapping), `bed_total_max_rent` (default `"N/A"`).

See [`examples/sample_deal.json`](examples/sample_deal.json) for a complete
example.

---

## 5. QC checklist (run after every deal)

- [ ] **Unit count** matches the source (and the Pres. RR total row).
- [ ] **Occupancy %** is reasonable (typically 90–98% for stabilized assets;
      flag < 88%).
- [ ] **No blank Market Rent** cells — every unit needs one (col U).
- [ ] **Model units** show `Model` in col Q (not counted as occupied by default).
- [ ] **Vacant units** show `0` contract rent and `0` LTL.
- [ ] **Recent Leases** rolling periods look sane (more activity in the longer
      windows). All-`n/a` usually means lease start dates are missing.
- [ ] **LTL** signs make sense (negative = contract above UW market).
- [ ] **Unit Mix totals** (units, SF) agree with the Pres. RR total row.
- [ ] The **audit block** ("Links from source data") totals match the summary.

---

## 6. Common issues & fixes

| Symptom | Likely cause / fix |
|---------|--------------------|
| Unit count off | Source had subtotal/group rows mistaken for units — re-parse skipping them. |
| All contract rents 0 | The in-place/lease-rent column wasn't found — point to the right source column. |
| Occupancy looks low | NTV/VL not recognized — they should map to occupied; extend `status_map` if the source uses an unusual label. |
| Recent Leases all `n/a` | Lease start dates missing — Move-In is used as the placeholder; long-tenured renewals won't appear as recent. |
| LTL columns blank | No UW market rent given for that floor plan — supply `uw_market_rents`. |
| A unit wrongly shows as model | Tell me the unit # and its real tenant/status. |

---

## 7. Iterating

Send a source file and I'll process it, return the exhibits, and note any
assumptions. We refine the intake mapping and config over a few test deals
until it's dialed in for your typical sources.
