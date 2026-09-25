"""
runner — Entry point: run all five reports and write CSVs to *output_dir*.

Usage
-----
    python -m nctrack.runner --data data/ --output /tmp/nctrack_out/
"""

from __future__ import annotations

import argparse
from pathlib import Path

from nctrack import alertes, mens_global, mens_lignes, pareto, rap_hebdo
from nctrack.config import LegacyConfig
from nctrack.loader import load


def run(data_dir: str | Path, output_dir: str | Path, cfg=None) -> None:
    """Generate all five output CSVs in *output_dir*."""
    if cfg is None:
        cfg = LegacyConfig()

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    ds = load(data_dir)

    rh_rows = rap_hebdo.compute(ds, cfg)
    (out / "rap_hebdo.csv").write_text(rap_hebdo.to_csv(rh_rows), encoding="utf-8")

    pt_rows = pareto.compute(ds, cfg)
    (out / "pareto.csv").write_text(pareto.to_csv(pt_rows), encoding="utf-8")

    al_rows = alertes.compute(ds, cfg)
    (out / "alertes.csv").write_text(alertes.to_csv(al_rows), encoding="utf-8")

    ml_rows = mens_lignes.compute(ds, cfg, al_rows)
    (out / "mens_lignes.csv").write_text(mens_lignes.to_csv(ml_rows), encoding="utf-8")

    mg_rows = mens_global.compute(ds, cfg, al_rows)
    (out / "mens_global.csv").write_text(mens_global.to_csv(mg_rows), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Run NCTrack reports")
    parser.add_argument("--data", default="data", help="Path to data/ directory")
    parser.add_argument("--output", default="sorties", help="Path to output directory")
    args = parser.parse_args()
    run(args.data, args.output)
    print(f"Reports written to {args.output}/")


if __name__ == "__main__":
    main()
