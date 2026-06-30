"""Command-line entry point.

    python -m rent_roll_processor.cli --units units.json --config config.json --out exhibits.xlsx

`units.json`  : list of normalized unit dicts (see schema.Unit / README).
`config.json` : deal config (see schema.RollConfig / README).

Both can be combined into one file with top-level keys ``config`` and
``units`` and passed via ``--deal deal.json``.
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import date

from .schema import Unit, RollConfig
from .workbook import build_exhibits


def _load_config(d: dict) -> RollConfig:
    return RollConfig(
        property_name=d["property_name"],
        as_of_date=d["as_of_date"],
        uw_market_rents=d.get("uw_market_rents", {}),
        reno_tiers=d.get("reno_tiers", {}),
        status_map=d.get("status_map"),
        bed_total_max_rent=d.get("bed_total_max_rent", "N/A"),
    )


def _load_units(rows: list, config: RollConfig) -> list:
    return [Unit.from_dict(r, status_map=config.status_map) for r in rows]


def build_from_deal(deal: dict, out_path: str) -> str:
    config = _load_config(deal["config"])
    units = _load_units(deal["units"], config)
    # Fill missing per-unit property names from the deal config.
    for u in units:
        if not u.property_name:
            u.property_name = config.property_name
    return build_exhibits(units, config, out_path)


def main(argv=None):
    p = argparse.ArgumentParser(description="Build the rent roll exhibits workbook.")
    p.add_argument("--deal", help="Single JSON with 'config' and 'units' keys.")
    p.add_argument("--units", help="JSON list of normalized unit dicts.")
    p.add_argument("--config", help="JSON deal config.")
    p.add_argument("--out", "-o", required=True, help="Output .xlsx path.")
    args = p.parse_args(argv)

    if args.deal:
        deal = json.load(open(args.deal))
    elif args.units and args.config:
        deal = {"config": json.load(open(args.config)),
                "units": json.load(open(args.units))}
    else:
        p.error("Provide --deal, or both --units and --config.")
        return 2

    out = build_from_deal(deal, args.out)
    print(f"Wrote {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
