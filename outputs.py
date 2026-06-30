"""Exports: CSVs, run manifest."""
from __future__ import annotations
import json
import datetime as dt
from pathlib import Path
import pandas as pd

from . import excel_report, persist, changes as changes_mod, engine as engine_mod, ingest as ingest_mod


def export(outdir: str, per_layer: pd.DataFrame, sw: pd.DataFrame,
           mr: pd.DataFrame, netting: pd.DataFrame, summary: pd.DataFrame,
           recon: pd.DataFrame, params) -> Path:
    out = Path(outdir)
    out.mkdir(parents=True, exist_ok=True)

    per_layer.drop(columns=[c for c in per_layer.columns if c.startswith("wb_")],
                   errors="ignore").to_csv(out / "per_layer.csv", index=False)
    sw.to_csv(out / "space_weather_by_manufacturer.csv", index=False)
    mr.to_csv(out / "max_risk_by_layer.csv", index=False)
    netting.to_csv(out / "netting_waterfalls.csv", index=False)
    summary.to_csv(out / "summary_grid.csv", index=False)
    recon.to_csv(out / "reconciliation.csv", index=False)

    # Q-on-Q: diff against the latest prior persisted run (before saving this one)
    prior_sel = params.raw.get("changes", {}).get("compare_to") \
        if isinstance(params.raw.get("changes"), dict) else None
    prior = persist.load_prior(str(params.as_at), explicit=prior_sel)
    chg = changes_mod.compute(per_layer, summary, prior)


    excluded = getattr(engine_mod.run_engine, "last_excluded", None)
    corrections = getattr(ingest_mod.load, "last_corrections", None)
    excel_report.write_results(
        str(out / f"Space_RDS_results_{params.as_at}.xlsx"),
        per_layer, sw, mr, summary, params,
        source=params.ingest.get("source", "?"), recon=recon,
        changes=chg, excluded=excluded,
        corrections=corrections)

    # persist this run so future quarters can diff against it
    persist.save_run(params, per_layer, summary)
    if chg:
        print(f"      changes vs {chg['prior_as_at']}: "
              f"{chg['layers']['summary']['new_layers']} new, "
              f"{chg['layers']['summary']['dropped_layers']} dropped")

    manifest = {
        "quarter": params.quarter,
        "as_at": str(params.as_at),
        "run_at": dt.datetime.now().isoformat(timespec="seconds"),
        "rows": len(per_layer),
        "layers": int(per_layer["layer_key"].nunique()),
        "reconciliation_ok": bool((recon["status"] == "OK").all()),
        "outputs": sorted(f.name for f in out.iterdir() if f.is_file()
                          and f.suffix in (".csv", ".xlsx")),
    }
    (out / "manifest.json").write_text(json.dumps(manifest, indent=2))
    return out
