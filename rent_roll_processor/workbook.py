"""Workbook writer.

Builds the 4-tab "Rent Roll Exhibits" workbook:

    1. Unit Mix & RR Summary       (per floor plan)
    2. Recent Leases               (per floor plan, rolling lease windows)
    3. Unit Mix based on # of Beds  (per bed count)
    4. Pres. Rent Roll             (one row per unit)

Layout, banding, theme colors and number formats were lifted cell-for-cell
from the Heritage Hill Estates reference exhibits and are reproduced here.
"""
from __future__ import annotations

from datetime import date
from typing import List, Optional, Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

from .schema import Unit, RollConfig
from . import aggregate as agg
from . import derive

# --- Theme colors (standard Office theme, matching the reference file) -------
C_BLUE = Color(theme=4, tint=-0.25)     # accent1 darker -> header band
C_TOTAL = Color(theme=8, tint=0.8)      # accent5 light  -> total/wtd avg rows
C_STAT = Color(theme=0, tint=-0.05)     # near-white     -> summary stat band
C_WHITE = Color(theme=0)
C_BLACK = Color(theme=1)
C_TITLE = Color(theme=3)                # dk2 -> title text
GRAY = "FF808080"

HEADER_FILL = PatternFill("solid", fgColor=C_BLUE)
TOTAL_FILL = PatternFill("solid", fgColor=C_TOTAL)
STAT_FILL = PatternFill("solid", fgColor=C_STAT)
VACANT_FILL = PatternFill("solid", fgColor="FFFFFF00")   # yellow (255,255,0)
VACANT_FONT = Font(size=10, color="FF0000FF")            # blue (0,0,255)

TITLE_FONT = Font(bold=True, size=15, color=C_TITLE)
PROP_FONT = Font(bold=True, size=15, color=GRAY)
SEC_FONT = Font(bold=True, size=12, color=C_WHITE)
HDR_FONT = Font(bold=True, size=11, color=C_WHITE)
HDR_FONT_SM = Font(bold=True, size=10, color=C_WHITE)
TOTAL_FONT = Font(bold=True, size=11, color=C_BLACK)
STAT_FONT = Font(bold=True, size=11, color=C_BLACK)
BODY_FONT = Font(size=11, color=C_BLACK)

THIN = Side(style="thin")
DOTTED = Side(style="dotted")
DASHDOT = Side(style="dashDot")
THICK = Side(style="thick")

# Number formats
NF_MONEY = '"$"#,##0_);[Red]\\("$"#,##0\\)'
NF_USD = '"$"#,##0'
NF_USD2 = '"$"#,##0.00'
NF_PCT = "0.00%"
NF_PCT1 = "0.0%"
NF_SF = '#,##0\\ "SF"'
NF_UNITS = '#,##0\\ "Units"'
NF_UNITS1 = '#\\ "Units"'        # Tot Units column
NF_OCCUNITS = '#0\\ "Units"'     # Occ Units column (always shows a digit)
NF_INT = "#,##0"
NF_DATE = "mm\\/dd\\/yyyy"
NF_DATE2 = "mm-dd-yy"


def _c(ws, coord, value=None, font=None, fill=None, nf=None,
       halign=None, valign=None, border=None):
    cell = ws[coord]
    if value is not None:
        cell.value = value
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if nf:
        cell.number_format = nf
    if halign or valign:
        cell.alignment = Alignment(horizontal=halign, vertical=valign)
    if border:
        cell.border = border
    return cell


def _band(ws, row, c_start, c_end, fill=HEADER_FILL):
    """Fill a contiguous horizontal band (used for header / total bands)."""
    for col in range(c_start, c_end + 1):
        ws.cell(row, col).fill = fill


def _band_font(ws, row, c_start, c_end, font, halign=None):
    """Apply a font (and optional alignment) across a band so that empty
    cells share the style.  Needed for centerContinuous banners to render."""
    for col in range(c_start, c_end + 1):
        cell = ws.cell(row, col)
        cell.font = font
        if halign:
            cell.alignment = Alignment(horizontal=halign, vertical="center")


def _set_widths(ws, widths):
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w


def _money(v):
    return float(v) if v is not None else None


# ---------------------------------------------------------------------------
# Pres. Rent Roll
# ---------------------------------------------------------------------------
PR_WIDTHS = {"A": 0.9, "D": 23.4, "E": 13.0, "F": 15.9, "G": 18.0, "H": 9.1,
             "I": 9.1, "J": 12.1, "K": 17.9, "L": 12.1, "M": 6.4, "O": 6.9,
             "P": 6.9, "Q": 6.4, "R": 5.9, "S": 11.0, "U": 11.4, "V": 11.4,
             "W": 11.4, "X": 10.9, "Y": 10.4, "Z": 12.9, "AA": 14.4,
             "AB": 15.4, "AC": 0.9, "AD": 0.9, "AE": 15.4, "AF": 10.4, "AG": 9.4}

# row 6 / row 7 two-line headers, keyed by column letter
PR_HDR = {
    "D": ("Property", "Name"), "E": ("Unit", "Number"), "F": ("Unit", "Type"),
    "G": ("Floor", "Plan"), "H": ("Unit", "Designation"), "I": ("", "Reno'd?"),
    "J": ("Reno", "Type"), "K": ("Tenant", "Name"), "L": ("Net", "Sq. Ft."),
    "M": ("#", "Bed"), "N": ("#", "Bath"), "O": ("Lease", "Type"),
    "P": ("Reno", "Stat"), "Q": ("Occ", "Stat*"), "R": ("", "MTM"),
    "S": ("Renew", "Status"), "T": ("Asking", "Rent"), "U": ("Market", "Rent"),
    "V": ("Contract", "Rent"), "W": ("Conc/", "Specl"), "X": ("Emp", "Disc"),
    "Y": ("Other", "Income"), "Z": ("Move In", "Date"),
    "AA": ("Lease", "Start Date"), "AB": ("Lease", "End Date"),
}
PR_UW_HDR = {"AE": "Market Rent", "AF": "LTL", "AG": "LTL %"}


def _build_pres_rent_roll(wb, units: List[Unit], config: RollConfig, totals):
    ws = wb.create_sheet("Pres. Rent Roll")
    ws.sheet_view.showGridLines = False
    _set_widths(ws, PR_WIDTHS)

    # Title
    _band_font(ws, 3, 3, 29, TITLE_FONT)
    _c(ws, "C3", "PRESENTATION RENT ROLL", font=TITLE_FONT, valign="center")
    for col in range(3, 29):  # thick underline across C3:AB3
        ws.cell(3, col).border = Border(bottom=THICK)
    _c(ws, "AD3", config.property_name, font=PROP_FONT, halign="right")

    # Section bands (row 5) - centerContinuous needs the style across the band
    _band(ws, 5, 3, 29)  # C..AC
    _band_font(ws, 5, 3, 29, SEC_FONT, halign="centerContinuous")
    ws["C5"].value = "UNIT INFORMATION"
    ws["O5"].value = "UNIT STATUS"
    ws["V5"].value = "CURRENT LEASE"

    # Column headers (rows 6 & 7)
    _band(ws, 6, 3, 29)
    _band(ws, 7, 3, 29)
    for letter, (l1, l2) in PR_HDR.items():
        ha = "left" if letter in ("D", "K") else "center"
        if l1:
            _c(ws, f"{letter}6", l1, font=HDR_FONT, halign=ha)
        # narrow columns use a smaller header font to fit (matches reference)
        f7 = HDR_FONT_SM if letter in ("H", "I", "J") else HDR_FONT
        _c(ws, f"{letter}7", l2, font=f7, halign=ha)
    for letter, txt in PR_UW_HDR.items():
        _c(ws, f"{letter}7", txt, font=Font(bold=True, size=11, color=C_BLACK), halign="center")

    # Data rows
    r = 8
    for u in units:
        _c(ws, f"D{r}", u.property_name or config.property_name, font=BODY_FONT, halign="left")
        _c(ws, f"E{r}", u.unit_id, font=BODY_FONT, halign="center")
        if u.unit_type is not None:
            _c(ws, f"F{r}", u.unit_type, font=BODY_FONT, halign="center")
        _c(ws, f"G{r}", u.floor_plan, font=BODY_FONT, halign="center")
        if u.designation:
            _c(ws, f"H{r}", u.designation, font=BODY_FONT, halign="center")
        if u.renovated:
            _c(ws, f"I{r}", u.renovated, font=BODY_FONT, halign="center")
        if u.reno_type:
            _c(ws, f"J{r}", u.reno_type, font=BODY_FONT, halign="center")
        _c(ws, f"K{r}", u.tenant_name or ("VACANT" if u.is_vacant else None),
           font=BODY_FONT, halign="left")
        _c(ws, f"L{r}", u.sqft, font=BODY_FONT, nf=NF_SF, halign="right")
        _c(ws, f"M{r}", u.beds, font=BODY_FONT, nf="0.#", halign="center")
        _c(ws, f"N{r}", u.baths, font=BODY_FONT, nf="0.#", halign="center")
        if u.lease_type:
            _c(ws, f"O{r}", u.lease_type, font=BODY_FONT, halign="center")
        if u.reno_status:
            _c(ws, f"P{r}", u.reno_status, font=BODY_FONT, halign="center")
        _c(ws, f"Q{r}", u.occupancy, font=BODY_FONT, nf=NF_INT, halign="center")
        if u.mtm:
            _c(ws, f"R{r}", u.mtm, font=BODY_FONT, halign="center")
        if u.renew_status:
            _c(ws, f"S{r}", u.renew_status, font=BODY_FONT, halign="center")
        if u.asking_rent is not None:
            _c(ws, f"T{r}", u.asking_rent, font=BODY_FONT, nf=NF_MONEY, halign="right")
        _c(ws, f"U{r}", _money(u.market_rent), font=BODY_FONT, nf=NF_MONEY, halign="right")
        _c(ws, f"V{r}", _money(u.contract_rent), font=BODY_FONT, nf=NF_MONEY, halign="right")
        _c(ws, f"W{r}", _money(u.concession), font=BODY_FONT, nf=NF_MONEY, halign="right")
        _c(ws, f"X{r}", _money(u.emp_discount), font=BODY_FONT, nf=NF_MONEY, halign="right")
        _c(ws, f"Y{r}", _money(u.other_income), font=BODY_FONT, nf=NF_MONEY, halign="right")
        if u.move_in:
            _c(ws, f"Z{r}", u.move_in, font=BODY_FONT, nf=NF_DATE, halign="right")
        if u.lease_start:
            _c(ws, f"AA{r}", u.lease_start, font=BODY_FONT, nf=NF_DATE, halign="right")
        if u.lease_end:
            _c(ws, f"AB{r}", u.lease_end, font=BODY_FONT, nf=NF_DATE, halign="right")
        uw, ltl, ltlpct = agg.loss_to_lease(u, config)
        if uw is not None:
            _c(ws, f"AE{r}", uw, font=BODY_FONT, nf=NF_USD, halign="center")
            _c(ws, f"AF{r}", ltl, font=BODY_FONT, nf=NF_USD, halign="center")
            _c(ws, f"AG{r}", ltlpct, font=BODY_FONT, nf=NF_PCT, halign="center")
        r += 1

    # Total / Wtd. Average row
    t = r
    top = Border(top=THIN)
    _c(ws, f"C{t}", "Total / Wtd. Average", font=TOTAL_FONT, nf=NF_DATE2, halign="left")
    _c(ws, f"G{t}", totals.tot_units, font=TOTAL_FONT, nf=NF_UNITS, halign="center", border=top)
    _c(ws, f"L{t}", totals.sqft, font=TOTAL_FONT, nf=NF_SF, halign="right", border=top)
    _c(ws, f"U{t}", totals.mkt_rent, font=TOTAL_FONT, nf=NF_MONEY, halign="right", border=top)
    _c(ws, f"V{t}", totals.cont_rent, font=TOTAL_FONT, nf=NF_MONEY, halign="right", border=top)
    _c(ws, f"W{t}", totals.conc, font=TOTAL_FONT, nf=NF_MONEY, halign="right", border=top)
    _c(ws, f"X{t}", totals.emp_disc, font=TOTAL_FONT, nf=NF_MONEY, halign="right", border=top)
    _c(ws, f"Y{t}", totals.other_inc, font=TOTAL_FONT, nf=NF_MONEY, halign="right", border=top)
    _c(ws, f"AE{t}", totals.uw_market_total, font=BODY_FONT, nf=NF_USD, halign="center")
    _c(ws, f"AF{t}", totals.ltl_total, font=BODY_FONT, nf=NF_USD, halign="center")
    _c(ws, f"AG{t}", totals.ltl_pct_total, font=BODY_FONT, nf=NF_PCT, halign="center")
    # thin top border for the remaining columns to draw a clean line
    for col in range(4, 34):
        cur = ws.cell(t, col)
        if cur.border.top.style is None:
            cur.border = top

    ws.freeze_panes = "A8"
    return ws


# ---------------------------------------------------------------------------
# Unit Mix & RR Summary  +  Bed Mix (share most layout)
# ---------------------------------------------------------------------------
MIX_WIDTHS = {"A": 0.9, "C": 0.9, "D": 29.6, "E": 28.0, "F": 15.4, "G": 14.1,
              "H": 15.4, "I": 13.9, "J": 12.1, "K": 12.4, "L": 14.4, "M": 12.4,
              "N": 15.0, "O": 13.1, "Q": 12.4, "R": 0.9, "S": 13.4, "T": 9.4,
              "U": 7.1, "V": 9.4, "W": 9.4}

MIX_HEADERS = [
    ("D", "Property Name", "left"), ("E", None, "left"),
    ("F", "Occ Units", "center"), ("G", "Tot Units", "center"),
    ("H", "Sq. Ft", "center"), ("I", "Avg SF", "center"),
    ("J", "Mkt Rent", "center"), ("K", "Avg Mkt/Unit", "center"),
    ("L", "AvgMkt/SF", "center"), ("M", "Cont Rent", "center"),
    ("N", "Avg Cont/Unit", "center"), ("O", "AvgCont/SF", "center"),
    ("P", "Max Rent", "center"), ("Q", "Other Inc.", "center"),
    ("S", "Sq. Ft. (Occ)", "center"),
]


def _write_mix_row(ws, r, prop, key, g: agg.GroupSummary, key_nf=None, fill=None, font=None):
    font = font or BODY_FONT
    def put(letter, value, nf, ha="right"):
        _c(ws, f"{letter}{r}", value, font=font, fill=fill, nf=nf, halign=ha, valign="center")
    put("D", prop, "General", "left")
    put("E", key, key_nf or "General", "left")
    put("F", g.occ_units, NF_OCCUNITS, "center")
    put("G", g.tot_units, NF_UNITS1)
    put("H", g.sqft, NF_SF)
    put("I", g.avg_sf, NF_SF)
    put("J", g.mkt_rent, NF_USD)
    put("K", g.avg_mkt_unit, NF_USD)
    put("L", g.avg_mkt_sf, NF_USD2)
    put("M", g.cont_rent, NF_USD)
    put("N", g.avg_cont_unit, NF_USD)
    put("O", g.avg_cont_sf, NF_USD2)
    put("P", g.max_rent, NF_USD)
    put("Q", g.other_inc, NF_USD)
    put("S", g.occ_sqft, NF_SF)


def _write_stat_block(ws, start_row, totals, label_col="C", val_col="G"):
    """Total Units / Occupancy / Annualized stat band (cols C..G)."""
    rows = [
        ("Total Units", totals.tot_units, NF_UNITS),
        ("Occupied Units", totals.occ_units, NF_UNITS),
        ("Unit Occupancy", totals.occ_units / totals.tot_units if totals.tot_units else 0, NF_PCT1),
        ("Total Square Feet", totals.sqft, NF_SF),
        ("Occupied Square Feet", totals.occ_sqft, NF_SF),
        ("Square Feet Occupancy", totals.occ_sqft / totals.sqft if totals.sqft else 0, NF_PCT1),
        (None, None, None),  # blank spacer
        ("Market Rent Annualized", totals.mkt_rent * 12, NF_USD),
        ("In-Place Rent Annualized", totals.cont_rent * 12, NF_USD),
        ("Other Income Annualized", totals.other_inc * 12, NF_USD),
    ]
    r = start_row
    for label, value, nf in rows:
        if label is None:
            r += 1
            continue
        _band(ws, r, 3, 7, STAT_FILL)  # C..G
        _c(ws, f"{label_col}{r}", label, font=STAT_FONT, halign="left")
        _c(ws, f"{val_col}{r}", value, font=STAT_FONT, nf=nf, halign="right")
        r += 1
    return r


def _write_audit_block(ws, start_row, totals, col="U", include_conc=True):
    """'Links from source data:' QC table."""
    val_col = get_column_letter(openpyxl.utils.column_index_from_string(col) + 2)
    rows = [
        ("Total # units", totals.tot_units, NF_INT),
        ("Occ Units", totals.occ_units, NF_INT),
        ("Unit %Occ", totals.occ_units / totals.tot_units if totals.tot_units else 0, NF_PCT1),
        ("Net Sq Ft", totals.sqft, NF_INT),
        ("SF Occ", totals.occ_sqft, NF_INT),
        ("% SF Occ", totals.occ_sqft / totals.sqft if totals.sqft else 0, NF_PCT1),
        ("Mkt Rent", totals.mkt_rent, NF_INT),
        ("Cont Rent", totals.cont_rent, NF_INT),
    ]
    if include_conc:
        rows.append(("Conc", totals.conc, NF_INT))
        rows.append(("Other Inc", totals.other_inc, NF_INT))
    _c(ws, f"{col}{start_row}", "Links from source data:", font=Font(bold=True, size=11, color=C_BLACK))
    r = start_row + 1
    for label, value, nf in rows:
        _c(ws, f"{col}{r}", label, font=BODY_FONT)
        _c(ws, f"{val_col}{r}", value, font=BODY_FONT, nf=nf, halign="right")
        r += 1


def _build_mix_sheet(wb, sheet_name, title, units, config, totals, groups,
                     key_header, key_nf, key_values, total_max_rent, audit=True):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    _set_widths(ws, MIX_WIDTHS)

    last_col = 17  # Q
    _band_font(ws, 3, 3, last_col, TITLE_FONT)
    _c(ws, "C3", title, font=TITLE_FONT)
    for col in range(3, last_col + 1):
        ws.cell(3, col).border = Border(bottom=THICK)
    _c(ws, "R3", config.property_name, font=PROP_FONT, halign="right")

    # Header row 5
    _band(ws, 5, 3, last_col, HEADER_FILL)
    for letter, text, ha in MIX_HEADERS:
        label = key_header if letter == "E" else text
        if label is not None:
            _c(ws, f"{letter}5", label, font=HDR_FONT, halign=ha, valign="center")

    # Data rows
    r = 6
    for g, kv in zip(groups, key_values):
        _write_mix_row(ws, r, config.property_name, kv, g, key_nf=key_nf)
        r += 1

    # Total / Wtd Average row
    tr = r
    _band(ws, tr, 3, last_col, TOTAL_FILL)
    ws.cell(tr, 19).fill = TOTAL_FILL  # S column too
    _c(ws, f"C{tr}", "Total/Wtd Average", font=TOTAL_FONT, halign="left", valign="center")
    tot_group = agg.GroupSummary(
        key="", occ_units=totals.occ_units, tot_units=totals.tot_units,
        sqft=totals.sqft, avg_sf=totals.avg_sf, mkt_rent=totals.mkt_rent,
        avg_mkt_unit=totals.avg_mkt_unit, avg_mkt_sf=totals.avg_mkt_sf,
        cont_rent=totals.cont_rent, avg_cont_unit=totals.avg_cont_unit,
        avg_cont_sf=totals.avg_cont_sf, max_rent=totals.max_rent,
        other_inc=totals.other_inc, occ_sqft=totals.occ_sqft)
    _write_mix_row(ws, tr, "", "", tot_group, fill=TOTAL_FILL, font=TOTAL_FONT)
    ws[f"F{tr}"].number_format = NF_UNITS1   # total Occ Units uses '#' not '#0'
    ws[f"C{tr}"].value = "Total/Wtd Average"  # restore label (D was blanked)
    # Max rent override (bed tab uses literal "N/A")
    if total_max_rent is not None:
        _c(ws, f"P{tr}", total_max_rent, font=TOTAL_FONT, nf=NF_USD,
           fill=TOTAL_FILL, halign="right", valign="center")
    # D/E of total row should be blank
    ws[f"D{tr}"].value = None
    ws[f"E{tr}"].value = None

    # Summary stat band
    stat_end = _write_stat_block(ws, tr + 2, totals)

    if audit:
        audit_row = max(41, stat_end + 4)
        _write_audit_block(ws, audit_row, totals, col="U", include_conc=True)
    return ws


def _build_unit_mix(wb, units, config, totals):
    groups = agg.unit_mix(units, config.model_occupied, config.net_concession)
    return _build_mix_sheet(
        wb, "Unit Mix & RR Summary", "UNIT MIX & RENT ROLL SUMMARY",
        units, config, totals, groups,
        key_header="Floor Plan", key_nf="General",
        key_values=[g.key for g in groups],
        total_max_rent=totals.max_rent, audit=True)


def _build_bed_mix(wb, units, config, totals):
    groups = agg.bed_mix(units, config.model_occupied, config.net_concession)
    widths = dict(MIX_WIDTHS); widths["D"] = 40.0; widths["E"] = 18.0
    key_values = []
    for g in groups:
        try:
            key_values.append(int(float(g.key)))
        except (ValueError, TypeError):
            key_values.append(g.key)
    ws = _build_mix_sheet(
        wb, "Unit Mix based on # of Beds",
        "UNIT MIX & RENT ROLL SUMMARY (Based on # of beds only)",
        units, config, totals, groups,
        key_header="Bed", key_nf='#0\\ "Bed/s"',
        key_values=key_values,
        total_max_rent=config.bed_total_max_rent, audit=False)
    _set_widths(ws, {"D": 40.0, "E": 18.0})
    return ws


# ---------------------------------------------------------------------------
# Recent Leases
# ---------------------------------------------------------------------------
RL_WIDTHS = {"A": 0.6, "C": 0.6, "D": 28.4, "E": 20.1, "F": 13.6, "G": 15.4,
             "H": 13.6, "I": 13.1, "J": 15.4, "K": 12.1, "L": 14.4, "N": 12.1,
             "O": 9.6, "P": 12.1, "Q": 9.6, "R": 12.1, "S": 9.6, "T": 12.1,
             "U": 9.6, "V": 12.1, "W": 9.6, "X": 0.6, "Y": 16.4, "Z": 17.4,
             "AA": 9.4}

RL_HEADERS = [
    ("D", "Property Name", "left"), ("E", "Floor Plan", "left"),
    ("F", "Occ Units", "center"), ("G", "Tot Units", "center"),
    ("H", "Sq. Ft", "center"), ("I", "Avg SF", "center"),
    ("J", "Market Rent", "center"), ("K", "Market Rent", "center"),
    ("L", "In-Place Rent", "center"), ("M", "% of Market", "center"),
    ("N", "180 Days", "center"), ("O", "# Leases", "center"),
    ("P", "120 Days", "center"), ("Q", "# Leases", "center"),
    ("R", "90 Days", "center"), ("S", "# Leases", "center"),
    ("T", "60 Days", "center"), ("U", "# Leases", "center"),
    ("V", "30 Days", "center"), ("W", "# Leases", "center"),
]
RL_WINDOW_COLS = {180: ("N", "O"), 120: ("P", "Q"), 90: ("R", "S"),
                  60: ("T", "U"), 30: ("V", "W")}


def _build_recent_leases(wb, units, config, totals):
    ws = wb.create_sheet("Recent Leases")
    ws.sheet_view.showGridLines = False
    _set_widths(ws, RL_WIDTHS)
    last_col = 23  # W

    _band_font(ws, 3, 3, last_col, TITLE_FONT)
    _c(ws, "C3", "RECENT LEASE SUMMARY", font=TITLE_FONT)
    for col in range(3, last_col + 1):
        ws.cell(3, col).border = Border(bottom=THICK)
    _c(ws, "W3", config.property_name, font=PROP_FONT, halign="right")
    # RR date box
    _c(ws, "Y3", config.as_of_date, font=Font(bold=True, size=11),
       nf=NF_DATE2, halign="center",
       border=Border(left=DASHDOT, right=DASHDOT, top=DASHDOT, bottom=DASHDOT))
    _c(ws, "Z3", "<---- RR Date", font=Font(size=11, color=C_BLACK))

    # Section labels row 5 (centerContinuous banners)
    _band(ws, 5, 3, last_col, HEADER_FILL)
    _band_font(ws, 5, 3, last_col, HDR_FONT, halign="centerContinuous")
    ws["C5"].value = "All Units"
    ws["K5"].value = "Occupied Units"
    ws["N5"].value = "Recent Leases"

    # Header row 6
    _band(ws, 6, 3, last_col, HEADER_FILL)
    for letter, text, ha in RL_HEADERS:
        nf = '#\\ "Days"' if text and text.endswith("Days") else None
        _c(ws, f"{letter}6", text, font=HDR_FONT, halign=ha, valign="center", nf=nf)

    rows = agg.recent_leases(units, config.as_of_date, config.model_occupied, config.net_concession)
    r = 7
    for row in rows:
        _c(ws, f"D{r}", config.property_name, font=BODY_FONT, halign="left")
        _c(ws, f"E{r}", row.key, font=BODY_FONT, halign="left")
        _c(ws, f"F{r}", row.occ_units, font=BODY_FONT, nf=NF_OCCUNITS, halign="right", valign="center")
        _c(ws, f"G{r}", row.tot_units, font=BODY_FONT, nf=NF_UNITS1, halign="right", valign="center")
        _c(ws, f"H{r}", row.sqft, font=BODY_FONT, nf=NF_SF, halign="right", valign="center")
        _c(ws, f"I{r}", row.avg_sf, font=BODY_FONT, nf=NF_SF, halign="right", valign="center")
        _c(ws, f"J{r}", row.avg_mkt_unit, font=BODY_FONT, nf=NF_USD, halign="right", valign="center")
        _c(ws, f"K{r}", row.occ_market_rent, font=BODY_FONT, nf=NF_USD, halign="right", valign="center")
        _c(ws, f"L{r}", row.in_place_rent, font=BODY_FONT, nf=NF_USD, halign="right", valign="center")
        _c(ws, f"M{r}", row.pct_of_market, font=BODY_FONT, nf=NF_PCT1, halign="right", valign="center")
        for days, (acol, ccol) in RL_WINDOW_COLS.items():
            w = row.windows[days]
            _c(ws, f"{acol}{r}", w["avg"], font=BODY_FONT, nf=NF_USD, halign="right", valign="center")
            _c(ws, f"{ccol}{r}", w["count"], font=BODY_FONT, halign="right", valign="center")
        r += 1

    # Total / Wtd Average
    tr = r
    _band(ws, tr, 3, last_col, TOTAL_FILL)
    _c(ws, f"C{tr}", "Total/Wtd Average", font=TOTAL_FONT, halign="left", valign="center")
    _c(ws, f"F{tr}", totals.occ_units, font=TOTAL_FONT, nf=NF_UNITS1, fill=TOTAL_FILL, halign="right", valign="center")
    _c(ws, f"G{tr}", totals.tot_units, font=TOTAL_FONT, nf=NF_UNITS1, fill=TOTAL_FILL, halign="right", valign="center")
    _c(ws, f"H{tr}", totals.sqft, font=TOTAL_FONT, nf=NF_SF, fill=TOTAL_FILL, halign="right", valign="center")
    _c(ws, f"I{tr}", totals.avg_sf, font=TOTAL_FONT, nf=NF_SF, fill=TOTAL_FILL, halign="right", valign="center")
    _c(ws, f"J{tr}", totals.avg_mkt_unit, font=TOTAL_FONT, nf=NF_USD, fill=TOTAL_FILL, halign="right", valign="center")
    _c(ws, f"K{tr}", totals.occ_market_rent, font=TOTAL_FONT, nf=NF_USD, fill=TOTAL_FILL, halign="right", valign="center")
    _c(ws, f"L{tr}", totals.in_place_rent, font=TOTAL_FONT, nf=NF_USD, fill=TOTAL_FILL, halign="right", valign="center")
    _c(ws, f"M{tr}", totals.pct_of_market, font=TOTAL_FONT, nf=NF_PCT1, fill=TOTAL_FILL, halign="right", valign="center")
    for days, (acol, ccol) in RL_WINDOW_COLS.items():
        _c(ws, f"{acol}{tr}", agg.NA, font=TOTAL_FONT, nf=NF_USD, fill=TOTAL_FILL, halign="right", valign="center")
        _c(ws, f"{ccol}{tr}", totals.window_counts[days], font=TOTAL_FONT, nf=NF_INT, fill=TOTAL_FILL, halign="right" if ccol != "W" else "center", valign="center")

    # dotted separators between the three sections
    for rr in range(5, tr + 1):
        for sep_col in ("K", "N"):
            cell = ws[f"{sep_col}{rr}"]
            b = cell.border
            cell.border = Border(left=DOTTED, right=b.right, top=b.top, bottom=b.bottom)

    _write_audit_block(ws, max(58, tr + 14), totals, col="Z", include_conc=False)
    return ws


# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# OneLineRR  (Checking table + one-line-per-unit data region)
# ---------------------------------------------------------------------------
YELLOW = PatternFill("solid", fgColor="FFFFFF00")
OLR_SM = Font(size=10)
OLR_SMB = Font(bold=True, size=10)
# Statuses the Checking formulas treat as occupied (matches the reference tool;
# non-revenue units are "in use" for this checking view).
_OCC_STATUSES = ("Occ", "Admin", "Down", "Super", "Model")

# Checking-table header: (col letter, title, halign)
OLR_CHECK_HDR = [
    ("C", "Floor Plan", "left"), ("D", "Unit Type", "left"), ("E", "BD", "center"),
    ("F", "BA", "center"), ("G", "Renovated", "center"), ("H", "Units", "right"),
    ("I", "# Occ", "right"), ("J", "#Vac", "right"), ("K", "Avg SF", "right"),
    ("L", "Total SF", "right"), ("M", "Occ SF", "right"), ("N", "Vac SF", "right"),
    ("O", "Market Rent ", "right"), ("P", "Rent", "right"), ("Q", "Other Income", "right"),
    ("R", "Concessions", "right"), ("S", "Emp Discounts", "right"),
]
# Data-region header: (col letter, title).  The Other Income / Concession /
# Employee-Discount detail columns (from AB onward) are generated dynamically
# per deal from each unit's line items, so they are not listed here.
OLR_DATA_HDR = [
    ("A", "Property"), ("B", "#"), ("C", "Occupancy"), ("D", "Floorplan"),
    ("E", "Bd"), ("F", "Ba"), ("G", "Renovated"), ("H", "Unit #"),
    ("I", "Unit Type"), ("J", "Unit Sq Ft"), ("K", "Resident/Tenant Name"),
    ("L", "Market Rent"), ("M", "Rent"), ("N", "Move in"), ("O", "Lease Start"),
    ("P", "Lease End"), ("Q", "Move Out"), ("R", "Other Income"),
    ("S", "Concessions"), ("T", "Employee Discounts"),
    ("V", "Unit Type (Manual)"), ("W", "Unit Type (Manual) - Updated"),
    ("Y", "RENT"),
]
OLR_OI_START = 28   # column AB: first Other Income detail column
OLR_WIDTHS = {"A": 21.6, "B": 4.1, "C": 24.0, "E": 12.6, "F": 4.6, "G": 9.4,
              "H": 9.6, "J": 11.6, "K": 18.6, "L": 17.6, "M": 11.1, "N": 11.6,
              "O": 11.6, "P": 11.4, "Q": 14.1, "R": 11.4, "S": 11.6, "T": 12.1,
              "V": 24.0, "Y": 12.1}


def _olr_unit_types(units):
    """Distinct unit types -> checking-table rows, sorted beds/baths/plan/reno.

    The Checking 'Floor Plan' label appends ', R' for renovated types so each
    row is unique (classic and reno of the same plan stay distinct)."""
    seen = {}
    sf_acc = {}
    for u in units:
        key = u.unit_type or u.floor_plan
        sf_acc.setdefault(key, []).append(u.sqft or 0)
        if key not in seen:
            reno = bool(u.renovated)
            label = u.floor_plan + (", R" if reno else "")
            seen[key] = {
                "code": key, "label": label, "beds": u.beds, "baths": u.baths,
                "reno": "Yes" if reno else "No",
            }
    INF = float("inf")
    for key, d in seen.items():
        rep_sf = sum(sf_acc[key]) / len(sf_acc[key])
        # priority: beds, baths, then unit SF (unknown beds/baths sort last)
        d["sort"] = (d["beds"] if d["beds"] is not None else INF,
                     d["baths"] if d["baths"] is not None else INF,
                     rep_sf, d["label"])
    return sorted(seen.values(), key=lambda d: d["sort"])


def _build_onelinerr(wb, units, config):
    ws = wb.create_sheet("OneLineRR")
    ws.sheet_view.showGridLines = False
    _set_widths(ws, OLR_WIDTHS)

    types = _olr_unit_types(units)
    n_types = len(types)

    # --- layout anchors -----------------------------------------------------
    CS = 5                       # first checking-table data row
    CE = CS + n_types - 1
    TOTALS = CE + 1
    PCT = TOTALS + 1
    NOTES = PCT + 4
    PROP_ROW = NOTES + 7
    DATE_ROW = PROP_ROW + 1
    GROUP_HDR = max(30, DATE_ROW + 2)
    DATA_HDR = GROUP_HDR + 1
    DS = DATA_HDR + 1            # first data row
    DE = DS + len(units) - 1

    # --- Checking table -----------------------------------------------------
    _c(ws, "C3", "Checking:", font=OLR_SMB)
    for letter, title, ha in OLR_CHECK_HDR:
        _c(ws, f"{letter}4", title, font=OLR_SMB, halign=ha)

    for i, t in enumerate(types):
        r = CS + i
        _c(ws, f"C{r}", t["label"], font=OLR_SM, fill=YELLOW, halign="left")
        _c(ws, f"D{r}", t["code"], font=OLR_SM, halign="left")
        _c(ws, f"E{r}", t["beds"], font=OLR_SM, fill=YELLOW, halign="center")
        _c(ws, f"F{r}", t["baths"], font=OLR_SM, fill=YELLOW, nf="0.0", halign="center")
        _c(ws, f"G{r}", t["reno"], font=OLR_SM, fill=YELLOW, halign="center")
        occ_count = "+".join(f'COUNTIFS($D${DS}:$D${DE},$C{r},$C${DS}:$C${DE},"{s}")'
                             for s in _OCC_STATUSES)
        occ_sf = "+".join(f'SUMIFS($J${DS}:$J${DE},$D${DS}:$D${DE},$C{r},$C${DS}:$C${DE},"{s}")'
                          for s in _OCC_STATUSES)
        fm = {
            "H": f"=COUNTIFS($D${DS}:$D${DE},$C{r})",
            "I": f"={occ_count}",
            "J": f'=COUNTIFS($D${DS}:$D${DE},$C{r},$C${DS}:$C${DE},"Vac")',
            "K": f'=IFERROR(AVERAGEIFS($J${DS}:$J${DE},$D${DS}:$D${DE},$C{r}),"n/a")',
            "L": f"=SUMIFS($J${DS}:$J${DE},$D${DS}:$D${DE},$C{r})",
            "M": f"={occ_sf}",
            "N": f'=SUMIFS($J${DS}:$J${DE},$D${DS}:$D${DE},$C{r},$C${DS}:$C${DE},"Vac")',
            "O": f"=SUMIFS($L${DS}:$L${DE},$D${DS}:$D${DE},$C{r})",
            "P": f"=SUMIFS($M${DS}:$M${DE},$D${DS}:$D${DE},$C{r})",
            "Q": f"=SUMIFS($R${DS}:$R${DE},$D${DS}:$D${DE},$C{r})",
            "R": f"=SUMIFS($S${DS}:$S${DE},$D${DS}:$D${DE},$C{r})",
            "S": f"=SUMIFS($T${DS}:$T${DE},$D${DS}:$D${DE},$C{r})",
        }
        for col, formula in fm.items():
            _c(ws, f"{col}{r}", formula, font=OLR_SM, nf=NF_INT, halign="right")

    # Totals row
    _c(ws, f"C{TOTALS}", "Totals", font=OLR_SMB, halign="left")
    for col in ("H", "I", "J", "L", "M", "N", "O", "P", "Q", "R", "S"):
        _c(ws, f"{col}{TOTALS}", f"=SUM({col}{CS}:{col}{CE})", font=OLR_SMB, nf=NF_INT, halign="right")
    # % row
    _c(ws, f"C{PCT}", "%", font=OLR_SMB, halign="left")
    _c(ws, f"I{PCT}", f"=I{TOTALS}/$H${TOTALS}", font=OLR_SMB, nf=NF_PCT1, halign="right")
    _c(ws, f"J{PCT}", f"=J{TOTALS}/$H${TOTALS}", font=OLR_SMB, nf=NF_PCT1, halign="right")
    _c(ws, f"M{PCT}", f'=IFERROR(M{TOTALS}/$L${TOTALS},"n/a")', font=OLR_SMB, nf=NF_PCT1, halign="right")
    _c(ws, f"N{PCT}", f'=IFERROR(N{TOTALS}/$L${TOTALS},"n/a")', font=OLR_SMB, nf=NF_PCT1, halign="right")

    # Notes
    _c(ws, f"C{NOTES}", "Notes:", font=OLR_SM)
    notes = [
        "1. Vacant units market rent is the per-unit market/asking rent from the source.",
        "2. BD/BA count is from the Unit Type mapping above (edit the yellow cells to adjust).",
    ]
    if any(u.occupancy == "Model" for u in units):
        notes.append("3. Model units carry market rent as a placeholder; excluded from revenue.")
    for k, txt in enumerate(notes, 1):
        _c(ws, f"C{NOTES + k}", txt, font=OLR_SM)

    # Property name / RR date
    _c(ws, f"C{PROP_ROW}", "Property Name", font=OLR_SM)
    _c(ws, f"D{PROP_ROW}", config.property_name, font=OLR_SM)
    _c(ws, f"C{DATE_ROW}", "Rent Roll Date", font=OLR_SM)
    _c(ws, f"D{DATE_ROW}", config.as_of_date, font=OLR_SMB, nf=NF_DATE2, halign="left")

    # --- Data region --------------------------------------------------------
    # Dynamic Other Income detail columns: one per distinct line item across
    # units (falls back to a single "Other Income" column if no per-item
    # breakdown was provided but a total exists).
    oi_names = []
    for u in units:
        for name in (u.other_income_items or {}):
            if name not in oi_names:
                oi_names.append(name)
    if not oi_names and any(u.other_income for u in units):
        oi_names = ["Other Income"]
    oi_cols = {name: OLR_OI_START + i for i, name in enumerate(oi_names)}

    has_conc = any(u.concession for u in units)
    conc_col = (max(oi_cols.values()) + 2) if oi_cols else OLR_OI_START
    L = get_column_letter

    _c(ws, f"Y{GROUP_HDR}", "Rent", font=OLR_SMB)
    if oi_cols:
        _c(ws, f"{L(OLR_OI_START)}{GROUP_HDR}", "Other Income", font=OLR_SMB)
    if has_conc:
        _c(ws, f"{L(conc_col)}{GROUP_HDR}", "Concessions", font=OLR_SMB)
    for letter, title in OLR_DATA_HDR:
        _c(ws, f"{letter}{DATA_HDR}", title, font=OLR_SMB)
    for name, col in oi_cols.items():
        _c(ws, f"{L(col)}{DATA_HDR}", name, font=OLR_SMB)
        ws.column_dimensions[L(col)].width = 13.0
    if has_conc:
        _c(ws, f"{L(conc_col)}{DATA_HDR}", "Concession", font=OLR_SMB)
        ws.column_dimensions[L(conc_col)].width = 13.0

    oi_first = L(OLR_OI_START)
    oi_last = L(max(oi_cols.values())) if oi_cols else oi_first
    conc_letter = L(conc_col)
    # rightmost used column, for the full-row vacant highlight
    row_end_col = max([25] + list(oi_cols.values()) + ([conc_col] if has_conc else []))

    for idx, u in enumerate(units):
        r = DS + idx
        _c(ws, f"A{r}", f"=$D${PROP_ROW}", font=OLR_SM, halign="left")
        _c(ws, f"B{r}", 1 if idx == 0 else f"=B{r-1}+1", font=OLR_SM, halign="left")
        _c(ws, f"C{r}", f'=IF(K{r}="Vacant","Vac",IF(K{r}="Model","Model",'
                        f'IF(K{r}="Admin","Admin",IF(K{r}="Down","Down",'
                        f'IF(K{r}="Super","Super","Occ")))))', font=OLR_SM, halign="left")
        _c(ws, f"D{r}", f"=INDEX($C${CS}:$C${CE},MATCH(I{r},$D${CS}:$D${CE},0))", font=OLR_SM, halign="left")
        _c(ws, f"E{r}", f"=INDEX($E${CS}:$E${CE},MATCH(D{r},$C${CS}:$C${CE},0))", font=OLR_SM, halign="left")
        _c(ws, f"F{r}", f"=INDEX($F${CS}:$F${CE},MATCH(D{r},$C${CS}:$C${CE},0))", font=OLR_SM, halign="left")
        _c(ws, f"G{r}", f"=INDEX($G${CS}:$G${CE},MATCH(D{r},$C${CS}:$C${CE},0))", font=OLR_SM, halign="left")
        _c(ws, f"H{r}", u.unit_id, font=OLR_SM, halign="left")
        _c(ws, f"I{r}", u.unit_type, font=OLR_SM, halign="left")
        _c(ws, f"J{r}", u.sqft, font=OLR_SM, nf='_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)')
        _c(ws, f"K{r}", u.tenant_name, font=OLR_SM, halign="left")
        _c(ws, f"L{r}", _money(u.market_rent), font=OLR_SM, nf="#,##0.00", halign="right")
        _c(ws, f"M{r}", f"=Y{r}", font=OLR_SM, nf="#,##0.00", halign="right")
        if u.move_in:
            _c(ws, f"N{r}", u.move_in, font=OLR_SM, nf=NF_DATE, halign="right")
        if u.lease_start:
            _c(ws, f"O{r}", u.lease_start, font=OLR_SM, nf=NF_DATE, halign="right")
        if u.lease_end:
            _c(ws, f"P{r}", u.lease_end, font=OLR_SM, nf=NF_DATE, halign="right")
        if getattr(u, "move_out", None):
            _c(ws, f"Q{r}", u.move_out, font=OLR_SM, nf=NF_DATE, halign="right")
        # Other Income totals over the dynamic detail columns.
        r_formula = f"=SUM({oi_first}{r}:{oi_last}{r})" if oi_cols else 0
        _c(ws, f"R{r}", r_formula, font=OLR_SM, nf="#,##0.00", halign="right")
        _c(ws, f"S{r}", f"={conc_letter}{r}" if has_conc else 0, font=OLR_SM, nf="#,##0.00", halign="right")
        _c(ws, f"T{r}", 0, font=OLR_SM, nf="#,##0.00", halign="right")
        _c(ws, f"Y{r}", _money(u.contract_rent), font=OLR_SM, nf="#,##0.00", halign="right")
        # Per-line-item Other Income values.
        items = u.other_income_items or {}
        for name, col in oi_cols.items():
            val = items.get(name) if items else (u.other_income if name == "Other Income" else 0)
            if val:
                _c(ws, f"{L(col)}{r}", _money(val), font=OLR_SM, nf="#,##0.00", halign="right")
        if has_conc and u.concession:
            _c(ws, f"{conc_letter}{r}", _money(u.concession), font=OLR_SM, nf="#,##0.00", halign="right")
        # Vacant units: highlight the whole data row (yellow fill, blue font).
        if u.occupancy == "Vac":
            for col in range(1, row_end_col + 1):
                cell = ws.cell(r, col)
                cell.fill = VACANT_FILL
                cell.font = VACANT_FONT

    ws.freeze_panes = f"A{DS}"
    return ws


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------
def build_exhibits(units: List[Unit], config: RollConfig, out_path: str) -> str:
    """Build the 5-tab exhibits workbook and save it to ``out_path``."""
    if not units:
        raise ValueError("No units to process.")

    # Deal-level derivations (vacant market rent, non-revenue offset).
    units = derive.apply_deal_rules(units, config)

    mix = agg.unit_mix(units, config.model_occupied, config.net_concession)
    totals = agg.compute_totals(units, config, mix=mix)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # drop default sheet

    # Tab order matches the reference exhibits file.
    _build_unit_mix(wb, units, config, totals)
    _build_recent_leases(wb, units, config, totals)
    _build_bed_mix(wb, units, config, totals)
    _build_pres_rent_roll(wb, units, config, totals)
    _build_onelinerr(wb, units, config)

    wb.save(out_path)
    return out_path
