"""Persist each run's layer set + headline grid so later runs can diff Q-on-Q."""
from __future__ import annotations
import json
from pathlib import Path
import pandas as pd

STORE = Path("runs")   # runs/<as_at>/{layers.csv,grid.csv,meta.json}


def save_run(params, per_layer: pd.DataFrame, grid: pd.DataFrame) -> Path:
    d = STORE / str(params.as_at)
    d.mkdir(parents=True, exist_ok=True)
    cols = ["program_id", "layer_id", "layer_key", "entity", "mapping_code",
            "spacecraft_id", "spacecraft_name", "orbit", "bus_manufacturer",
            "inception", "off_risk_date", "per_sc", "net_of_xol"]
    cols = [c for c in cols if c in per_layer.columns]
    per_layer[cols].to_csv(d / "layers.csv", index=False)
    grid.to_csv(d / "grid.csv", index=False)
    # S3123 grid — persisted when the S3123 path ran this quarter
    s3cfg = (params.raw.get("s3123_rds") or {})
    if s3cfg.get("enabled"):
        from . import s3123 as s3123_mod
        s3grid = s3123_mod.s3123_grid(per_layer, params)
        s3grid.to_csv(d / "s3123_grid.csv", index=False)
    (d / "meta.json").write_text(json.dumps(
        {"quarter": params.quarter, "as_at": str(params.as_at),
         "layers": int(per_layer["layer_key"].nunique())}, indent=2))
    return d


def list_runs() -> list[str]:
    if not STORE.exists():
        return []
    return sorted(p.name for p in STORE.iterdir() if (p / "layers.csv").exists())


def load_prior(current_as_at: str, explicit: str | None = None):
    """Return (layers_df, grid_df, as_at) for the chosen prior run, or None.

    explicit: a specific runs/<as_at> name; else the latest run before current.
    """
    runs = list_runs()
    if explicit:
        target = explicit if explicit in runs else None
    else:
        earlier = [r for r in runs if r < str(current_as_at)]
        target = earlier[-1] if earlier else None
    if not target:
        return None
    d = STORE / target
    s3path = d / "s3123_grid.csv"
    s3grid = pd.read_csv(s3path) if s3path.exists() else None
    return (pd.read_csv(d / "layers.csv"), pd.read_csv(d / "grid.csv"),
            target, s3grid)
