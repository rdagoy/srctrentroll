"""Output-file naming convention.

Every processed deal is named:

    Birgo RR Exhibits - <Property Name>v<mm.dd.yy>.xlsx

The `v` stands for "version" and is followed by the file-modified date
(today's date), which serves as the version stamp.  Example:

    Birgo RR Exhibits - Station J Townv07.01.26.xlsx
"""
from __future__ import annotations

from datetime import date
from typing import Optional


def output_filename(property_name: str, on_date: Optional[date] = None) -> str:
    """Return the standard exhibits filename for a deal.

    Args:
        property_name: the deal / property name.
        on_date: version date; defaults to today (the file-modified date).
    """
    stamp = (on_date or date.today()).strftime("%m.%d.%y")
    return f"Birgo RR Exhibits - {property_name}v{stamp}.xlsx"
