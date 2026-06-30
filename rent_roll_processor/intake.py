"""Generic source-file intake helpers.

These cover the mechanical part of Step 1 (read a spreadsheet/CSV into raw row
dicts, then rename columns to the normalized schema). The *judgement* part —
deciding which source column maps to which normalized field, where the header
row is, and which rows are real units — is done per source format.

Typical use::

    rows = read_table("source.xlsx", sheet="Rent Roll", header_row=5)
    units_raw = map_rows(rows, {
        "Unit": "unit_id", "Floorplan": "floor_plan", "SqFt": "sqft",
        "Beds": "beds", "Baths": "baths", "Market": "market_rent",
        "Rent": "contract_rent", "Status": "occupancy",
        "Lease From": "lease_start", "Lease To": "lease_end",
    })
    units = [Unit.from_dict(r) for r in units_raw if r.get("unit_id")]
"""
from __future__ import annotations

import csv
import os
from typing import List, Dict, Any, Optional


def read_table(path: str, sheet: Optional[str] = None,
               header_row: Optional[int] = None) -> List[Dict[str, Any]]:
    """Read a CSV/XLSX/XLS into a list of {column_header: value} dicts.

    Args:
        path: source file path.
        sheet: worksheet name (Excel only; defaults to the active/first sheet).
        header_row: 1-indexed row holding column names. If None, the first
            non-empty row is used as the header.
    """
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return _read_csv(path, header_row)
    if ext in (".xlsx", ".xlsm", ".xls"):
        return _read_excel(path, sheet, header_row)
    raise ValueError(f"Unsupported source extension: {ext}")


def _read_csv(path: str, header_row: Optional[int]) -> List[Dict[str, Any]]:
    with open(path, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    return _rows_to_dicts(rows, header_row)


def _read_excel(path: str, sheet: Optional[str],
                header_row: Optional[int]) -> List[Dict[str, Any]]:
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    return _rows_to_dicts(rows, header_row)


def _rows_to_dicts(rows: List[list], header_row: Optional[int]) -> List[Dict[str, Any]]:
    if not rows:
        return []
    if header_row is None:
        header_idx = next((i for i, r in enumerate(rows)
                           if any(c not in (None, "") for c in r)), 0)
    else:
        header_idx = header_row - 1
    headers = [str(c).strip() if c is not None else f"col{j}"
               for j, c in enumerate(rows[header_idx])]
    out = []
    for r in rows[header_idx + 1:]:
        if all(c in (None, "") for c in r):
            continue
        out.append({headers[j]: r[j] if j < len(r) else None
                    for j in range(len(headers))})
    return out


def map_rows(rows: List[Dict[str, Any]],
             column_map: Dict[str, str],
             constants: Optional[Dict[str, Any]] = None) -> List[Dict[str, Any]]:
    """Rename raw columns to normalized field names.

    Args:
        rows: raw dicts from ``read_table``.
        column_map: {source_header: normalized_field}.
        constants: fields applied to every row (e.g. property_name).
    """
    constants = constants or {}
    out = []
    for r in rows:
        d = dict(constants)
        for src, dst in column_map.items():
            if src in r:
                d[dst] = r[src]
        out.append(d)
    return out
