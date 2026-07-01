"""Unit Type Mapping — the two-step workflow for sources without bed/bath.

Step 1 (scan): call ``write_unit_type_mapping(units, path)`` to emit a small
Excel table listing every distinct unit type with its SF and unit count, plus
blank (yellow) Floor Plan / BD / BA columns for the analyst to fill.

Step 2 (incorporate): the analyst fills and returns the table; call
``read_unit_type_mapping(path)`` to get a dict and pass it as
``RollConfig.unit_type_map`` so the full model groups by the real floor plans.
"""
from __future__ import annotations

from collections import OrderedDict
from typing import Dict, List

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

_YELLOW = PatternFill("solid", fgColor="FFFFFF00")
_HDR = Font(bold=True, size=11)
_TITLE = Font(bold=True, size=14)
_THIN = Side(style="thin")

# (column letter, header, width)
_COLS = [("B", "Unit Type", 22), ("C", "Sq. Ft", 14), ("D", "# Units", 10),
         ("E", "Floor Plan", 26), ("F", "BD", 7), ("G", "BA", 7)]
_EDITABLE = ("E", "F", "G")   # analyst fills these (yellow)


def _distinct_types(units):
    """OrderedDict unit_type -> {sqfts:set, count:int}, sorted by SF then type."""
    agg: "OrderedDict[str, dict]" = OrderedDict()
    for u in units:
        key = u.unit_type or u.floor_plan
        d = agg.setdefault(key, {"sqfts": [], "count": 0})
        d["sqfts"].append(u.sqft or 0)
        d["count"] += 1
    items = sorted(agg.items(),
                   key=lambda kv: (sum(kv[1]["sqfts"]) / len(kv[1]["sqfts"]), str(kv[0])))
    return OrderedDict(items)


def _sf_label(sqfts) -> str:
    lo, hi = min(sqfts), max(sqfts)
    return f"{lo:g}" if lo == hi else f"{lo:g}–{hi:g}"


def write_unit_type_mapping(units, path: str, property_name: str = "") -> str:
    """Write the fill-in Unit Type Mapping table and return ``path``."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Unit Type Mapping"
    ws.sheet_view.showGridLines = False

    ws["B2"] = f"Unit Type Mapping — {property_name}".strip(" —")
    ws["B2"].font = _TITLE
    ws["B3"] = ("Fill the yellow columns (Floor Plan, BD, BA) for each unit type, "
                "then send the file back. Leave a cell blank to keep it unknown.")
    ws["B3"].font = Font(italic=True, size=10)

    for letter, title, width in _COLS:
        c = ws[f"{letter}5"]
        c.value = title
        c.font = _HDR
        c.alignment = Alignment(horizontal="left" if letter in ("B", "E") else "center")
        c.border = Border(bottom=_THIN)
        ws.column_dimensions[letter].width = width

    r = 6
    for utype, d in _distinct_types(units).items():
        ws[f"B{r}"] = utype
        ws[f"C{r}"] = _sf_label(d["sqfts"])
        ws[f"C{r}"].alignment = Alignment(horizontal="center")
        ws[f"D{r}"] = d["count"]
        ws[f"D{r}"].alignment = Alignment(horizontal="center")
        for letter in _EDITABLE:
            ws[f"{letter}{r}"].fill = _YELLOW
            ws[f"{letter}{r}"].alignment = Alignment(
                horizontal="left" if letter == "E" else "center")
        r += 1

    wb.save(path)
    return path


def read_unit_type_mapping(path: str) -> Dict[str, dict]:
    """Read a filled mapping table -> {unit_type: {floor_plan, beds, baths}}."""
    ws = openpyxl.load_workbook(path, data_only=True).active
    out: Dict[str, dict] = {}
    r = 6
    while ws[f"B{r}"].value not in (None, ""):
        utype = str(ws[f"B{r}"].value).strip()
        fp = ws[f"E{r}"].value
        beds = ws[f"F{r}"].value
        baths = ws[f"G{r}"].value
        out[utype] = {
            "floor_plan": str(fp).strip() if fp not in (None, "") else None,
            "beds": beds if beds not in (None, "") else None,
            "baths": baths if baths not in (None, "") else None,
        }
        r += 1
    return out
