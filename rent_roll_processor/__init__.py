"""Rent Roll Processor.

Turns a normalized list of units + a deal config into the standardized
4-tab "Rent Roll Exhibits" Excel workbook used for underwriting.

Public API:
    from rent_roll_processor import Unit, RollConfig, build_exhibits
"""
from .schema import Unit, RollConfig
from .workbook import build_exhibits
from . import intake

__all__ = ["Unit", "RollConfig", "build_exhibits", "intake"]
__version__ = "1.0.0"
