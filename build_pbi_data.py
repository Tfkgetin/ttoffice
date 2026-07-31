#!/usr/bin/env python3
"""
Self-refreshing Power BI data prep for the Space RDS.

Run from the repo root AFTER the pipeline has produced one or more quarterly
runs:

    python build_pbi_data.py            # discovers every output/<as_at>/ run
    python build_pbi_data.py --out pbi_data

For each quarter it reads output/<as_at>/per_layer.csv (+ manifest.json for the
quarter label), recomputes the Lloyd's syndicate shares on the FILED basis
(FUL+FIID, FIBL excluded — via src.s3123, so it ties JJ to the dollar), adds
clash / net-cascade / time dimensions, stacks all quarters, and writes a small
star schema under ./pbi_data:

    fact_per_layer.csv          layer x quarter   (the fact table)
    dim_spacecraft.csv          one row per spacecraft (latest attributes)
    dim_quarter.csv             quarter x as-at + sort order (time axis)
    fact_spacecraft_quarter.csv spacecraft x quarter + movement vs prior Q
    dim_scenario_params.csv     scenario reference card
    dim_risk_appetite.csv       appetite thresholds for RAG flags

Nothing is hard-coded to a quarter: drop a new output/<as_at>/ run in and re-run.
"""
from __future__ import annotations
import argparse, json, sys
from pathlib import Path
import pandas as pd

# --- optional import of the pipeline's own share logic (best: ties filed basis)
try:
    from src import s3123 as _s3123
    from src.parameters import Params
    _HAVE_SRC = True
except Exception:                                   # run outside the repo
    _HAVE_SRC = False


# ---------------------------------------------------------------- discovery ---
def discover_runs(root: Path):
    """Every output/<as_at>/per_layer.csv, oldest first."""
    runs = []
    for d in sorted(root.glob("*/")):
        pl = d / "per_layer.csv"
        if not pl.exists():
            continue
        as_at = d.name
        quarter = as_at
        mf = d / "manifest.json"
        if mf.exists():
            try:
                quarter = json.loads(mf.read_text()).get("quarter", as_at)
            except Exception:
                pass
        runs.append({"as_at": as_at, "quarter": quarter, "path": pl})
    return runs


def load_params(quarter: str):
    """config/<quarter>.yaml -> Params, or None (share recompute then falls back)."""
    if not _HAVE_SRC:
        return None
    for cand in (f"config/{quarter}.yaml", f"config/{quarter.replace(' ', '')}.yaml"):
        if Path(cand).exists():
            try:
                return Params.load(cand)
            except Exception as e:
                print(f"    ! could not load {cand}: {e}")
    return None


# ---------------------------------------------------------------- enrichment --
NUM = ["per_sc", "rpf", "altitude_km", "s3123_factor", "s2126_factor",
       "ext_qs", "igr_qs_ceded", "net_of_qs", "xol_ceded", "net_of_xol",
       "s3123_qs", "equity_usd", "total_fibl_ceded",
       "pf_fihl", "gd_fihl", "sd_fihl", "sw_fihl", "mr_fihl",
       "pf_ful", "gd_ful", "sd_ful", "pf_fiid", "gd_fiid", "sd_fiid"]

KEEP = ["program_id", "layer_id", "layer_key", "entity", "mapping_code",
        "spacecraft_id", "spacecraft_name", "orbit", "bus_manufacturer",
        "bus_family", "coverage", "placing_basis", "controlling_body",
        "is_consortium", "inception", "expiry", "on_risk_date", "off_risk_date"]


def altitude_band(orbit, km):
    if pd.isna(km) or km == 0:
        return {"LEO": "LEO (unbanded)", "MEO": "MEO", "GEO-GSO": "GEO"}.get(str(orbit), str(orbit))
    if km < 600:   return "LEO 400-600"
    if km < 800:   return "LEO 600-800"
    if km < 1200:  return "LEO 800-1200"
    if km < 1600:  return "LEO/MEO 1200-1600"
    if km < 35000: return "MEO"
    return "GEO"


def syndicate_share(df, params, factor_col):
    """FILED basis share (FUL+FIID, FIBL excluded). Uses src.s3123 when available."""
    if _HAVE_SRC and params is not None and factor_col in df.columns:
        try:
            return _s3123.s3123_share(df.copy(), _s3123._cfg(params), factor_col=factor_col)
        except Exception as e:
            print(f"    ! s3123_share fallback ({factor_col}): {e}")
    # fallback: per_sc x factor, consortium & non-FIBL only
    if factor_col not in df.columns:
        return pd.Series(0.0, index=df.index)
    isc = (df["is_consortium"].astype(str).str.lower().isin(["true", "1", "1.0"])
           if "is_consortium" in df.columns else df[factor_col] > 0)
    return (df["per_sc"] * df[factor_col]).where(isc & ~df["entity"].eq("FIBL"), 0.0)


def enrich(path, quarter, as_at, params):
    d = pd.read_csv(path, low_memory=False)
    for c in KEEP:
        if c not in d.columns:
            d[c] = pd.NA
    for c in NUM:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0) if c in d.columns else 0.0
    d["entity"] = d["entity"].astype(str)
    d["orbit"] = d["orbit"].astype(str).replace("nan", "")
    d["quarter"] = quarter
    d["as_at"] = as_at
    d["is_leo"] = (d["orbit"] == "LEO").astype(int)
    d["altitude_band"] = [altitude_band(o, k) for o, k in zip(d["orbit"], d["altitude_km"])]
    d["s3123_share"] = syndicate_share(d, params, "s3123_factor").values
    d["s2126_share"] = syndicate_share(d, params, "s2126_factor").values
    keep = KEEP + NUM + ["quarter", "as_at", "is_leo", "altitude_band",
                         "s3123_share", "s2126_share"]
    return d[[c for c in keep if c in d.columns]].copy()


# ------------------------------------------------------------------- movement -
def movement(fact, quarters):
    """spacecraft x quarter with status vs the immediately-prior quarter."""
    def prog(df):
        return (df.groupby("spacecraft_id")["program_id"]
                  .apply(lambda s: "|".join(sorted(map(str, set(s))))))
    rows = []
    for i, q in enumerate(quarters):
        cur = fact[fact["quarter"] == q]
        cur_e = cur.groupby(["spacecraft_id", "spacecraft_name", "orbit"])["per_sc"].sum().reset_index()
        cur_p = prog(cur)
        if i == 0:
            prev_e, prev_p = {}, {}
        else:
            pv = fact[fact["quarter"] == quarters[i - 1]]
            prev_e = pv.groupby("spacecraft_id")["per_sc"].sum().to_dict()
            prev_p = prog(pv).to_dict()
        for _, r in cur_e.iterrows():
            sid = r["spacecraft_id"]
            a = float(prev_e.get(sid, 0)); b = float(r["per_sc"])
            if i == 0:                     st = "Baseline"
            elif a == 0:                   st = "New"
            elif prev_p.get(sid) != cur_p.get(sid): st = "Renewed"
            else:                          st = "Continued"
            rows.append(dict(quarter=q, spacecraft_id=sid, spacecraft_name=r["spacecraft_name"],
                             orbit=r["orbit"], exposure=b, exposure_prev=a, delta=b - a,
                             movement_status=st))
        # non-renewals: in prior, absent now
        gone = set(prev_e) - set(cur_e["spacecraft_id"])
        for sid in gone:
            nm = fact.loc[fact["spacecraft_id"] == sid, "spacecraft_name"]
            rows.append(dict(quarter=q, spacecraft_id=sid,
                             spacecraft_name=(nm.iloc[0] if len(nm) else str(sid)),
                             orbit="", exposure=0.0, exposure_prev=float(prev_e[sid]),
                             delta=-float(prev_e[sid]), movement_status="Non-renewed"))
    return pd.DataFrame(rows)


# ----------------------------------------------------------------------- main -
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--output-root", default="output", help="dir holding <as_at>/ runs")
    ap.add_argument("--out", default="pbi_data", help="where to write the model")
    args = ap.parse_args()

    root = Path(args.output_root)
    runs = discover_runs(root)
    if not runs:
        sys.exit(f"No runs found under {root}/<as_at>/per_layer.csv — run the pipeline first.")

    print(f"Discovered {len(runs)} quarter(s): " + ", ".join(r['quarter'] for r in runs))
    frames = []
    for r in runs:
        params = load_params(r["quarter"])
        f = enrich(r["path"], r["quarter"], r["as_at"], params)
        frames.append(f)
        print(f"  {r['quarter']:10} rows={len(f):4}  gross={f['per_sc'].sum():,.0f}"
              f"  S3123={f['s3123_share'].sum():,.0f}  S2126={f['s2126_share'].sum():,.0f}")
    fact = pd.concat(frames, ignore_index=True)

    quarters = [r["quarter"] for r in runs]
    out = Path(args.out); out.mkdir(parents=True, exist_ok=True)
    fact.to_csv(out / "fact_per_layer.csv", index=False)

    # dim_quarter (time axis)
    pd.DataFrame({"quarter": quarters,
                  "as_at": [r["as_at"] for r in runs],
                  "sort_order": range(len(quarters))}).to_csv(out / "dim_quarter.csv", index=False)

    # dim_spacecraft (latest attributes)
    last = fact[fact["quarter"] == quarters[-1]]
    dim = (fact.sort_values("as_at").groupby("spacecraft_id")
             .agg(spacecraft_name=("spacecraft_name", "last"),
                  orbit=("orbit", "last"),
                  bus_manufacturer=("bus_manufacturer", "last"),
                  altitude_band=("altitude_band", "last")).reset_index())
    dim.to_csv(out / "dim_spacecraft.csv", index=False)

    # fact_spacecraft_quarter (movement)
    movement(fact, quarters).to_csv(out / "fact_spacecraft_quarter.csv", index=False)

    # dim_scenario_params
    pd.DataFrame([
        dict(scenario="Proton Flare",  loss_factor="0.05", orbit_scope="GEO-GSO", uses_rpf=False),
        dict(scenario="Generic Defect", loss_factor="0.50", orbit_scope="GEO-GSO, MEO", uses_rpf=True),
        dict(scenario="Space Debris",  loss_factor="LEO 0.40 / MEO 0.10 / GEO 0.05", orbit_scope="All", uses_rpf=False),
        dict(scenario="Space Weather", loss_factor="1.00", orbit_scope="Worst bus manufacturer", uses_rpf=False),
        dict(scenario="Max Risk",      loss_factor="1.00", orbit_scope="Largest spacecraft", uses_rpf=False),
    ]).to_csv(out / "dim_scenario_params.csv", index=False)

    # dim_risk_appetite (RAG thresholds) — Lloyd's read from config when available
    appetite = 50_000_000
    p_last = load_params(quarters[-1])
    if p_last is not None:
        appetite = ((p_last.raw.get("s3123_rds") or {}).get("lloyds_summary") or {}).get(
            "risk_appetite_usd", appetite)
    pd.DataFrame([
        dict(book="S3123 (Lloyd's)", appetite_usd=appetite),
        dict(book="IG",              appetite_usd=pd.NA),      # set your IG appetite here
    ]).to_csv(out / "dim_risk_appetite.csv", index=False)

    print(f"\nWrote {out}/ : fact_per_layer, dim_spacecraft, dim_quarter, "
          f"fact_spacecraft_quarter, dim_scenario_params, dim_risk_appetite")
    print("Power BI: relate fact_per_layer[spacecraft_id]->dim_spacecraft, "
          "fact_*[quarter]->dim_quarter (both single-direction).")


if __name__ == "__main__":
    main()
