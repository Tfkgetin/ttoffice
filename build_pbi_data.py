#!/usr/bin/env python3
"""
Self-refreshing Power BI STAR SCHEMA for the Space RDS.

Run from the repo root after the pipeline has produced one or more quarters:

    python run_space_rds.py          # -> output/<as_at>/
    python build_pbi_data.py         # -> pbi_data/

Discovers every output/<as_at>/ run and emits a tidy star schema. The key design
choice is LONG fact tables (one row per quarter x entity x scenario, and per
layer x scenario) rather than wide columns — that is what lets a single visual
serve every scenario/entity/perspective via field parameters, instead of one
visual per combination.

    FACTS
      fact_layer.csv            layer x quarter               detail, exposure, net cascade
      fact_layer_scenario.csv   layer x qtr x scen x entity   contribution ($)
      fact_scenario.csv         qtr x entity x scenario       gross / net / ceded
      fact_movement.csv         spacecraft x quarter          movement buckets
    DIMENSIONS
      dim_quarter.csv  dim_spacecraft.csv  dim_entity.csv
      dim_scenario.csv  dim_risk_appetite.csv

Nothing is hard-coded to a quarter: drop in a new output/<as_at>/ and re-run.
"""
from __future__ import annotations
import argparse, json, re, sys
from pathlib import Path
import pandas as pd

try:                                    # pipeline's own logic (preferred)
    from src import s3123 as _s3123
    from src.parameters import Params
    _HAVE_SRC = True
except Exception:
    _HAVE_SRC = False

SCEN = {"pf": "Proton Flare", "gd": "Generic Defect", "sd": "Space Debris",
        "sw": "Space Weather", "mr": "Max Risk"}
ENT = {"fihl": "FIHL", "ful": "FUL", "fiid": "FIID", "fibl": "FIBL"}
ENT_TYPE = {"FIHL": "Group", "FUL": "Operating", "FIID": "Operating",
            "FIBL": "Internal RI receiver", "S3123": "Syndicate", "S2126": "Syndicate"}
ENT_SORT = {"FIHL": 1, "FUL": 2, "FIID": 3, "FIBL": 4, "S3123": 5, "S2126": 6}
SCEN_SORT = {"Proton Flare": 1, "Generic Defect": 2, "Space Debris": 3,
             "Space Weather": 4, "Max Risk": 5}


# --------------------------------------------------------------- discovery ---
def discover(root: Path):
    runs = []
    for d in sorted(p for p in root.iterdir() if p.is_dir()):
        if not (d / "per_layer.csv").exists():
            continue
        q = d.name
        mf = d / "manifest.json"
        if mf.exists():
            try:
                q = json.loads(mf.read_text()).get("quarter", d.name)
            except Exception:
                pass
        runs.append({"as_at": d.name, "quarter": q, "dir": d})
    return runs


def load_params(quarter: str):
    if not _HAVE_SRC:
        return None
    for c in (f"config/{quarter}.yaml", f"config/{quarter.replace(' ', '')}.yaml"):
        if Path(c).exists():
            try:
                return Params.load(c)
            except Exception as e:
                print(f"    ! {c}: {e}")
    return None


# -------------------------------------------------------------- fact_layer ---
NUMC = ["per_sc", "rpf", "altitude_km", "s3123_factor", "s2126_factor",
        "ext_qs", "igr_qs_ceded", "net_of_qs", "xol_ceded", "net_of_xol",
        "s3123_qs", "equity_usd", "total_fibl_ceded",
        # scenario bases: debris_dr drives Space Debris, net_ext is Max Risk's gross
        "debris_dr", "net_ext"]
DIMC = ["layer_key", "program_id", "layer_id", "entity", "mapping_code", "spacecraft_id",
        "spacecraft_name", "orbit", "bus_manufacturer", "bus_family", "coverage",
        "placing_basis", "controlling_body", "is_consortium",
        "inception", "expiry", "on_risk_date", "off_risk_date"]


def band(orbit, km):
    try:
        km = float(km)
    except Exception:
        km = 0
    if not km:
        return {"LEO": "LEO (no altitude)", "MEO": "MEO",
                "GEO-GSO": "GEO"}.get(str(orbit), str(orbit) or "Unknown")
    if km < 600:   return "LEO 400-600"
    if km < 800:   return "LEO 600-800"
    if km < 1200:  return "LEO 800-1200"
    if km < 1600:  return "LEO/MEO 1200-1600"
    if km < 35000: return "MEO"
    return "GEO"


def share(df, params, col):
    """Lloyd's share on the FILED basis (FUL+FIID, FIBL excluded)."""
    if col not in df.columns:
        return pd.Series(0.0, index=df.index)
    if _HAVE_SRC and params is not None:
        try:
            return _s3123.s3123_share(df.copy(), _s3123._cfg(params), factor_col=col)
        except Exception as e:
            print(f"    ! s3123_share fallback ({col}): {e}")
    isc = (df["is_consortium"].astype(str).str.lower().isin(["true", "1", "1.0"])
           if "is_consortium" in df.columns else df[col] > 0)
    return (df["per_sc"] * df[col]).where(isc & ~df["entity"].eq("FIBL"), 0.0)


def build_layer(run, params):
    d = pd.read_csv(run["dir"] / "per_layer.csv", low_memory=False)
    for c in DIMC:
        if c not in d.columns:
            d[c] = pd.NA
    for c in NUMC:
        d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0) if c in d.columns else 0.0
    d["entity"] = d["entity"].astype(str)
    d["orbit"] = d["orbit"].astype(str).replace("nan", "")
    d["quarter"] = run["quarter"]
    d["altitude_band"] = [band(o, k) for o, k in zip(d["orbit"], d["altitude_km"])]
    d["s3123_share"] = share(d, params, "s3123_factor").values
    d["s2126_share"] = share(d, params, "s2126_factor").values
    d["ceded_total"] = d["per_sc"] - d["net_of_xol"]
    # Executive-summary axes: inception vintage, RPF band, coarse asset class.
    inc = pd.to_datetime(d["inception"], errors="coerce")
    d["inception_month"] = inc.dt.strftime("%Y-%m").fillna("Unknown")
    d["inception_month_sort"] = (inc.dt.year * 100 + inc.dt.month).fillna(0).astype(int)
    d["rpf_band"] = d["rpf"].map(lambda v: f"RPF {v:.2f}" if v > 0 else "No RPF")
    d["orbit_class"] = d["orbit"].map(
        {"GEO-GSO": "GEO", "GEO": "GEO", "MEO": "MEO", "LEO": "LEO"}).fillna("Other")
    d["layer_uid"] = (d["quarter"].astype(str) + "|" + d["program_id"].astype(str)
                      + "|" + d["layer_id"].astype(str))
    cols = (["layer_uid", "quarter"] + DIMC + NUMC
            + ["altitude_band", "s3123_share", "s2126_share", "ceded_total",
               "inception_month", "inception_month_sort", "rpf_band", "orbit_class"])
    return d[[c for c in cols if c in d.columns]].copy(), d


# ----------------------------------------------- fact_layer_scenario (LONG) ---
def build_layer_scenario(raw, quarter):
    """Unpivot the engine's per-layer scenario contributions into long form."""
    rows = []
    id_cols = [c for c in ["layer_uid", "spacecraft_id", "spacecraft_name", "orbit",
                           "bus_manufacturer", "altitude_band", "orbit_class",
                           "rpf_band", "inception_month"] if c in raw.columns]
    for col in raw.columns:
        m = re.fullmatch(r"(pf|gd|sd|sw|mr)_(fihl|ful|fiid|fibl)(?:_(direct|ceded))?", col)
        if m:
            slug, e, sub = m.groups()
            ent, scen = ENT[e], SCEN[slug]
            persp = "Gross" if sub in (None, "direct") else "Ceded in"
        else:
            m2 = re.fullmatch(r"(s3123|s2126)_(pf|gd|sd|sw|mr)_(g|n)", col)
            if not m2:
                continue
            syn, slug, gn = m2.groups()
            ent, scen = syn.upper(), SCEN[slug]
            persp = "Gross" if gn == "g" else "Net"
        v = pd.to_numeric(raw[col], errors="coerce").fillna(0)
        if v.abs().sum() == 0:
            continue
        blk = raw.loc[v != 0, id_cols].copy()
        blk["quarter"] = quarter
        blk["entity"] = ent
        blk["scenario"] = scen
        blk["perspective"] = persp
        blk["contribution"] = v[v != 0].values
        rows.append(blk)
    return pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()


# ---------------------------------------------------- fact_scenario (LONG) ----
def build_scenario(run, params, per_layer_raw):
    """quarter x entity x scenario -> gross / net, from the pipeline's own grids."""
    out = []
    sg = run["dir"] / "summary_grid.csv"
    if sg.exists():
        g = pd.read_csv(sg)

        def pick(r, base):
            if r.get("entity") in ("FIHL", "FIBL"):
                for alt in (f"{base}_incl_addons", f"{base}_receiver"):
                    if alt in g.columns and pd.notna(r.get(alt)):
                        return float(r[alt])
            return float(r.get(base, 0) or 0)

        for _, r in g.iterrows():
            out.append(dict(quarter=run["quarter"], entity=r.get("entity"),
                            scenario=r.get("scenario"), detail=r.get("detail"),
                            gross=pick(r, "gross"), net=pick(r, "net"),
                            gross_ex_addons=float(r.get("gross", 0) or 0),
                            net_ex_addons=float(r.get("net", 0) or 0)))
    s3 = run["dir"] / "s3123_grid.csv"
    if s3.exists():
        for _, r in pd.read_csv(s3).iterrows():
            out.append(dict(quarter=run["quarter"], entity=r.get("entity", "S3123"),
                            scenario=r.get("scenario"), detail=r.get("detail"),
                            gross=float(r.get("gross", 0) or 0),
                            net=float(r.get("net", 0) or 0),
                            gross_ex_addons=float(r.get("gross", 0) or 0),
                            net_ex_addons=float(r.get("net", 0) or 0)))
    # S2126 grid isn't exported as CSV — compute it when src is importable
    if _HAVE_SRC and params is not None and per_layer_raw is not None:
        try:
            s2 = _s3123.s3123_grid(per_layer_raw, params, cfg_key="s2126_rds",
                                   factor_col="s2126_factor", label="S2126")
            for _, r in s2.iterrows():
                out.append(dict(quarter=run["quarter"], entity="S2126",
                                scenario=r.get("scenario"), detail=r.get("detail"),
                                gross=float(r.get("gross", 0) or 0),
                                net=float(r.get("net", 0) or 0),
                                gross_ex_addons=float(r.get("gross", 0) or 0),
                                net_ex_addons=float(r.get("net", 0) or 0)))
        except Exception as e:
            print(f"    ! S2126 grid skipped: {e}")
    df = pd.DataFrame(out)
    if len(df):
        df["ceded"] = df["gross"] - df["net"]
    return df


# ------------------------------------------------------------ fact_movement ---
def build_movement(fact_layer, quarters):
    def progs(d):
        return (d.groupby("spacecraft_id")["program_id"]
                 .apply(lambda s: "|".join(sorted(map(str, set(s))))).to_dict())
    rows = []
    for i, q in enumerate(quarters):
        cur = fact_layer[fact_layer["quarter"] == q]
        ce = cur.groupby(["spacecraft_id", "spacecraft_name"])["per_sc"].sum().reset_index()
        cp = progs(cur)
        if i == 0:
            pe, pp = {}, {}
        else:
            pv = fact_layer[fact_layer["quarter"] == quarters[i - 1]]
            pe = pv.groupby("spacecraft_id")["per_sc"].sum().to_dict()
            pp = progs(pv)
        for _, r in ce.iterrows():
            sid, b = r["spacecraft_id"], float(r["per_sc"])
            a = float(pe.get(sid, 0))
            st = ("Baseline" if i == 0 else "New" if a == 0
                  else "Renewed" if pp.get(sid) != cp.get(sid) else "Continued")
            rows.append(dict(quarter=q, spacecraft_id=sid,
                             spacecraft_name=r["spacecraft_name"], exposure=b,
                             exposure_prior=a, delta=b - a, movement_status=st))
        for sid in set(pe) - set(ce["spacecraft_id"]):
            nm = fact_layer.loc[fact_layer["spacecraft_id"] == sid, "spacecraft_name"]
            rows.append(dict(quarter=q, spacecraft_id=sid,
                             spacecraft_name=(nm.iloc[0] if len(nm) else str(sid)),
                             exposure=0.0, exposure_prior=float(pe[sid]),
                             delta=-float(pe[sid]), movement_status="Non-renewed"))
    return pd.DataFrame(rows)


# ---------------------------------------------------------------------- main --
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--output-root", default="output")
    ap.add_argument("--out", default="pbi_data")
    a = ap.parse_args()

    root = Path(a.output_root)
    if not root.exists():
        sys.exit(f"{root}/ not found — run the pipeline first.")
    runs = discover(root)
    if not runs:
        sys.exit(f"No <as_at>/per_layer.csv under {root}/ — run the pipeline first.")
    print(f"Quarters found: {', '.join(r['quarter'] for r in runs)}")

    L, LS, S = [], [], []
    for r in runs:
        p = load_params(r["quarter"])
        lay, raw = build_layer(r, p)
        L.append(lay)
        ls = build_layer_scenario(raw, r["quarter"])
        if len(ls):
            LS.append(ls)
        sc = build_scenario(r, p, raw)
        if len(sc):
            S.append(sc)
        print(f"  {r['quarter']:>9}  layers={len(lay):4}  gross={lay['per_sc'].sum():>15,.0f}"
              f"  S3123={lay['s3123_share'].sum():>13,.0f}"
              f"  S2126={lay['s2126_share'].sum():>12,.0f}")

    fact_layer = pd.concat(L, ignore_index=True)
    quarters = [r["quarter"] for r in runs]
    out = Path(a.out); out.mkdir(parents=True, exist_ok=True)

    fact_layer.to_csv(out / "fact_layer.csv", index=False)
    (pd.concat(LS, ignore_index=True) if LS else pd.DataFrame()).to_csv(
        out / "fact_layer_scenario.csv", index=False)
    (pd.concat(S, ignore_index=True) if S else pd.DataFrame()).to_csv(
        out / "fact_scenario.csv", index=False)
    build_movement(fact_layer, quarters).to_csv(out / "fact_movement.csv", index=False)

    # ---- dimensions ----
    pd.DataFrame([dict(quarter=r["quarter"], as_at=r["as_at"], quarter_sort=i,
                       year=str(r["as_at"])[:4])
                  for i, r in enumerate(runs)]).to_csv(out / "dim_quarter.csv", index=False)

    (fact_layer.sort_values("quarter")
        .groupby("spacecraft_id")
        .agg(spacecraft_name=("spacecraft_name", "last"), orbit=("orbit", "last"),
             bus_manufacturer=("bus_manufacturer", "last"),
             altitude_band=("altitude_band", "last"),
             orbit_class=("orbit_class", "last"))
        .reset_index()).to_csv(out / "dim_spacecraft.csv", index=False)

    # Layer-grain conformed dimensions. fact_layer and fact_layer_scenario have no
    # relationship to each other, so a chart mixing a measure from one with an axis
    # from the other repeats the same total in every cell. These give both facts a
    # shared axis to hang off.
    fls = pd.concat(LS, ignore_index=True) if LS else pd.DataFrame(columns=["inception_month"])
    vint = sorted(set(fact_layer["inception_month"].dropna())
                  | set(fls.get("inception_month", pd.Series(dtype=str)).dropna()))
    pd.DataFrame([dict(inception_month=v,
                       inception_month_sort=(int(v[:4]) * 100 + int(v[5:7]))
                       if v != "Unknown" else 0,
                       year=v[:4] if v != "Unknown" else "Unknown")
                  for v in vint]).to_csv(out / "dim_vintage.csv", index=False)

    bands = sorted(set(fact_layer["rpf_band"].dropna())
                   | set(fls.get("rpf_band", pd.Series(dtype=str)).dropna()))
    pd.DataFrame([dict(rpf_band=b,
                       rpf_band_sort=0 if b == "No RPF" else int(round(float(b[4:]) * 100)))
                  for b in bands]).to_csv(out / "dim_rpf_band.csv", index=False)

    ents = sorted(set(fact_layer["entity"].dropna()) | set(ENT_SORT))
    pd.DataFrame([dict(entity=e, entity_type=ENT_TYPE.get(e, "Other"),
                       entity_sort=ENT_SORT.get(e, 99)) for e in ents]
                 ).to_csv(out / "dim_entity.csv", index=False)

    pd.DataFrame([
        dict(scenario="Proton Flare",   loss_factor="0.05", orbit_scope="GEO-GSO",
             uses_rpf=False, selection="Additive",
             method_desc="5% of sum insured on every GEO-GSO spacecraft. "
                         "No risk-period factor — driver is pure GEO book size."),
        dict(scenario="Generic Defect", loss_factor="0.50", orbit_scope="GEO-GSO, MEO",
             uses_rpf=True,  selection="Additive",
             method_desc="50% of sum insured x risk-period factor on GEO-GSO and MEO. "
                         "Driver is the RPF profile — how close each risk is to launch."),
        dict(scenario="Space Debris",   loss_factor="LEO .40 / MEO .10 / GEO .05",
             orbit_scope="All", uses_rpf=False, selection="Additive by orbit",
             method_desc="Orbit-tiered factors: 40% LEO, 10% MEO, 5% GEO. "
                         "Driver is the orbit mix, dominated by the LEO book."),
        dict(scenario="Space Weather",  loss_factor="1.00", orbit_scope="All",
             uses_rpf=False, selection="Worst bus manufacturer",
             method_desc="Total loss of every spacecraft built by the single worst "
                         "bus-manufacturer group. Selection, not additive — driver is "
                         "manufacturer concentration."),
        dict(scenario="Max Risk",       loss_factor="1.00", orbit_scope="All",
             uses_rpf=False, selection="Largest single spacecraft",
             method_desc="Total loss of the single largest spacecraft exposure. "
                         "Selection, not additive — driver is the peak single risk."),
    ]).assign(scenario_sort=lambda d: d["scenario"].map(SCEN_SORT)
              ).to_csv(out / "dim_scenario.csv", index=False)

    appetite = 50_000_000
    pl = load_params(quarters[-1])
    if pl is not None:
        appetite = (((pl.raw.get("s3123_rds") or {}).get("lloyds_summary") or {})
                    .get("risk_appetite_usd", appetite))
    pd.DataFrame([dict(entity="S3123", appetite_usd=appetite),
                  dict(entity="S2126", appetite_usd=appetite),
                  dict(entity="FIHL",  appetite_usd=None)]
                 ).to_csv(out / "dim_risk_appetite.csv", index=False)

    print(f"\n{out}/ written — 4 facts + 7 dimensions.")
    print("Relationships: fact_*[quarter]->dim_quarter, fact_*[spacecraft_id]->dim_spacecraft,")
    print("               fact_scenario/fact_layer_scenario[entity]->dim_entity,")
    print("               fact_scenario/fact_layer_scenario[scenario]->dim_scenario,")
    print("               fact_layer/fact_layer_scenario[inception_month]->dim_vintage,")
    print("               fact_layer/fact_layer_scenario[rpf_band]->dim_rpf_band.")


if __name__ == "__main__":
    main()
