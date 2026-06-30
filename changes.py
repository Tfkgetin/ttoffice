"""Q-on-Q changes: current run vs a prior persisted run."""
from __future__ import annotations
import pandas as pd

SCEN_ORDER = ["Proton Flare", "Space Weather", "Generic Defect",
              "Space Debris", "Max Risk"]


def layer_changes(cur: pd.DataFrame, prior: pd.DataFrame) -> dict:
    """New / dropped / moved layers between two runs (keyed on layer_key)."""
    cur = cur.copy(); prior = prior.copy()
    cur["layer_key"] = cur["layer_key"].astype(str)
    prior["layer_key"] = prior["layer_key"].astype(str)
    ck = set(cur["layer_key"]); pk = set(prior["layer_key"])

    new_keys = ck - pk
    dropped_keys = pk - ck
    common = ck & pk

    new = cur[cur["layer_key"].isin(new_keys)].copy()
    dropped = prior[prior["layer_key"].isin(dropped_keys)].copy()

    # exposure moves on layers present in both
    c = cur[cur["layer_key"].isin(common)].set_index("layer_key")["per_sc"]
    p = prior[prior["layer_key"].isin(common)].set_index("layer_key")["per_sc"]
    moved = pd.DataFrame({"prior": p, "current": c})
    moved["delta"] = moved["current"] - moved["prior"]
    moved = moved[moved["delta"].abs() > 1].copy()
    if len(moved):
        meta = cur.set_index("layer_key")[
            ["spacecraft_name", "entity", "bus_manufacturer"]]
        moved = moved.join(meta).sort_values("delta", key=abs, ascending=False)

    return {
        "new": new, "dropped": dropped, "moved": moved.reset_index(),
        "summary": {
            "new_layers": len(new), "dropped_layers": len(dropped),
            "moved_layers": len(moved),
            "new_exposure": float(new["per_sc"].sum()) if len(new) else 0.0,
            "dropped_exposure": float(dropped["per_sc"].sum()) if len(dropped) else 0.0,
            "net_move": float(moved["delta"].sum()) if len(moved) else 0.0,
        },
    }


def scenario_changes(cur_grid: pd.DataFrame, prior_grid: pd.DataFrame) -> pd.DataFrame:
    """Gross & net deltas per entity × scenario.

    Covers all four entities — FIHL (group), FUL & FIID (operating), and FIBL
    (internal RI receiver) — so the Changes matrix and the by-scenario net/gross
    tables show the complete entity picture.
    """
    def keyed(g, field):
        return g.set_index([g["entity"], g["scenario"]])[field]
    rows = []
    for entity in ["FIHL", "FUL", "FIBL", "FIID"]:
        for scen in SCEN_ORDER:
            cg = cur_grid[(cur_grid["entity"] == entity) &
                          (cur_grid["scenario"] == scen)]
            pg = prior_grid[(prior_grid["entity"] == entity) &
                            (prior_grid["scenario"] == scen)]
            if not len(cg) and not len(pg):
                continue
            def val(df, f):
                if len(df) and f in df and pd.notna(df.iloc[0][f]):
                    return float(df.iloc[0][f])
                return None
            cgr, pgr = val(cg, "gross"), val(pg, "gross")
            cnet, pnet = val(cg, "net"), val(pg, "net")
            rows.append({
                "entity": entity, "scenario": scen,
                "cur_gross": cgr, "prior_gross": pgr,
                "d_gross": (cgr - pgr) if None not in (cgr, pgr) else None,
                "d_gross_pct": ((cgr - pgr) / pgr) if (pgr) else None,
                "cur_net": cnet, "prior_net": pnet,
                "d_net": (cnet - pnet) if None not in (cnet, pnet) else None,
            })
    return pd.DataFrame(rows)


def compute(cur_layers, cur_grid, prior):
    """prior = (layers_df, grid_df, as_at) from persist.load_prior, or None."""
    if prior is None:
        return None
    p_layers, p_grid, p_asat = prior
    return {
        "prior_as_at": p_asat,
        "layers": layer_changes(cur_layers, p_layers),
        "scenarios": scenario_changes(cur_grid, p_grid),
    }
