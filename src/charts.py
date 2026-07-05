"""Chart generation (matplotlib PNGs) for embedding in the results workbook.

Senior-facing visuals in the house cool-navy palette. Requires matplotlib;
callers (outputs.py) skip gracefully if it is unavailable.

generate_all() returns {key: path} consumed by excel_report:
    exec_ranking     – group scenario loss profile (Exec Summary embed)
    scenario_compare – gross vs net by entity, binding scenario (Charts tab)
    waterfall        – netting cascade, worst scenario (Charts tab)
    exposure_bridge  – prior→current in-force walk (Charts tab)
    concentration    – top bus manufacturers by exposure (Charts tab)
"""
from __future__ import annotations
from pathlib import Path
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

# House cool-navy palette (matches excel_report tokens, '#'-prefixed for mpl)
INK = "#1F2933"; NAVY = "#1B3A5C"; NAVY_LT = "#5B86B0"
GREEN = "#2D6A3C"; AMBER = "#9A6410"; RED = "#C0392B"
SOFT = "#6B7785"; RULE = "#DCE3EA"; CREAM = "#EEF3F8"
SCEN_ORDER = ["Proton Flare", "Space Weather", "Generic Defect",
              "Space Debris", "Max Risk"]
ENT_ORDER = ["FIHL", "FUL", "FIBL", "FIID"]

plt.rcParams.update({
    "font.family": "DejaVu Sans", "font.size": 9,
    "axes.edgecolor": RULE, "axes.linewidth": 0.8,
    "axes.spines.top": False, "axes.spines.right": False,
    "axes.titlesize": 10, "axes.titleweight": "bold", "axes.titlecolor": NAVY,
    "text.color": INK, "axes.labelcolor": INK,
    "xtick.color": SOFT, "ytick.color": SOFT,
    "figure.dpi": 150,
})
_m = FuncFormatter(lambda x, _: f"{x/1e6:,.0f}")


def _save(fig, path):
    fig.savefig(path, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    return str(path)


def _val(row, f):
    v = row.get(f)
    try:
        return float(v) if v is not None and v == v else 0.0
    except (TypeError, ValueError):
        return 0.0


def exec_ranking(grid, outdir, entity="FUL"):
    """Scenario loss profile for one entity — gross vs net retained, ranked."""
    block = grid[grid["entity"] == entity].copy()
    if not len(block):
        return None
    block["_o"] = block["scenario"].map({s: i for i, s in enumerate(SCEN_ORDER)})
    block = block.sort_values("net", ascending=True)
    scen = block["scenario"].tolist()
    gross = [_val(r, "gross") for _, r in block.iterrows()]
    net = [_val(r, "net") for _, r in block.iterrows()]
    fig, ax = plt.subplots(figsize=(5.2, 3.0))
    y = range(len(scen))
    ax.barh(y, gross, color=RULE, height=0.62, label="Gross", zorder=2)
    ax.barh(y, net, color=NAVY, height=0.62, label="Net retained", zorder=3)
    ax.set_yticks(list(y)); ax.set_yticklabels(scen)
    ax.xaxis.set_major_formatter(_m)
    ax.set_xlabel("$m"); ax.set_title(f"{entity} — gross vs net retained by scenario")
    ax.legend(frameon=False, fontsize=8, loc="lower right")
    ax.grid(axis="x", color=RULE, linewidth=0.6, zorder=0)
    return _save(fig, Path(outdir) / "exec_ranking.png")


def scenario_compare(grid, outdir):
    """Gross loss by entity across scenarios — grouped bars (group risk shape)."""
    ents = [e for e in ENT_ORDER if (grid["entity"] == e).any()]
    if not ents:
        return None
    scens = [s for s in SCEN_ORDER if (grid["scenario"] == s).any()]
    import numpy as np
    x = np.arange(len(scens)); w = 0.8 / max(len(ents), 1)
    cols = {"FIHL": NAVY, "FUL": NAVY_LT, "FIBL": GREEN, "FIID": AMBER}
    fig, ax = plt.subplots(figsize=(6.4, 3.2))
    for i, e in enumerate(ents):
        vals = []
        for s in scens:
            row = grid[(grid["entity"] == e) & (grid["scenario"] == s)]
            vals.append(_val(row.iloc[0], "gross") if len(row) else 0.0)
        ax.bar(x + i * w - 0.4 + w / 2, vals, w, label=e,
               color=cols.get(e, SOFT), zorder=3)
    ax.set_xticks(x); ax.set_xticklabels(scens, rotation=15, ha="right", fontsize=8)
    ax.yaxis.set_major_formatter(_m); ax.set_ylabel("$m gross")
    ax.set_title("Gross loss by entity & scenario")
    ax.legend(frameon=False, fontsize=8, ncol=len(ents))
    ax.grid(axis="y", color=RULE, linewidth=0.6, zorder=0)
    return _save(fig, Path(outdir) / "scenario_compare.png")


def waterfall(grid, outdir, entity="FUL"):
    """Netting cascade gross→net for an entity's worst-gross scenario."""
    block = grid[grid["entity"] == entity]
    if not len(block):
        return None
    row = block.loc[block["gross"].idxmax()]
    gross = _val(row, "gross")
    steps = [("Ext QS", -_val(row, "ext_qs")),
             ("IGR QS", -_val(row, "igr_qs_ceded")),
             ("IGR XoL", -_val(row, "xol_ceded"))]
    labels = ["Gross"] + [s[0] for s in steps] + ["Net"]
    fig, ax = plt.subplots(figsize=(5.2, 3.2))
    run = gross
    ax.bar(0, gross, color=NAVY, zorder=3)
    for i, (lab, delta) in enumerate(steps, start=1):
        bottom = run + delta
        ax.bar(i, -delta if delta < 0 else delta, bottom=min(run, bottom),
               color=NAVY_LT, zorder=3)
        run = bottom
    ax.bar(len(steps) + 1, run, color=GREEN, zorder=3)
    ax.set_xticks(range(len(labels))); ax.set_xticklabels(labels, fontsize=8)
    ax.yaxis.set_major_formatter(_m); ax.set_ylabel("$m")
    ax.set_title(f"{entity} netting cascade — {row['scenario']}")
    ax.grid(axis="y", color=RULE, linewidth=0.6, zorder=0)
    return _save(fig, Path(outdir) / "waterfall.png")


def exposure_bridge(changes, closing_total, outdir):
    """Prior→current in-force walk as a bridge chart."""
    if not changes or not changes.get("layers"):
        return None
    s = changes["layers"]["summary"]
    new_exp = s["new_exposure"]; dropped = s["dropped_exposure"]; move = s["net_move"]
    opening = closing_total - new_exp + dropped - move
    labels = ["Opening", "+ New", "− Run-off", "± Reval", "Closing"]
    deltas = [new_exp, -dropped, move]
    cols = [GREEN if d >= 0 else RED for d in deltas]
    fig, ax = plt.subplots(figsize=(5.6, 3.2))
    ax.bar(0, opening, color=NAVY, zorder=3)
    run = opening
    for i, d in enumerate(deltas, start=1):
        ax.bar(i, abs(d), bottom=min(run, run + d), color=cols[i - 1], zorder=3)
        run += d
    ax.bar(4, run, color=NAVY, zorder=3)
    ax.set_xticks(range(5)); ax.set_xticklabels(labels, fontsize=8)
    ax.yaxis.set_major_formatter(_m); ax.set_ylabel("$m on-risk")
    pa = changes.get("prior_as_at", "prior")
    ax.set_title(f"Exposure bridge — {pa} → current")
    ax.grid(axis="y", color=RULE, linewidth=0.6, zorder=0)
    return _save(fig, Path(outdir) / "exposure_bridge.png")


def concentration(sw, outdir, n=8):
    """Top bus manufacturers by FIHL (group) exposure — correlation risk."""
    if sw is None or not len(sw):
        return None
    col = "sw_fihl" if "sw_fihl" in sw.columns else sw.columns[1]
    top = sw.sort_values(col, ascending=False).head(n).iloc[::-1]
    fig, ax = plt.subplots(figsize=(5.8, 3.2))
    ax.barh(range(len(top)), top[col], color=NAVY, zorder=3)
    ax.set_yticks(range(len(top)))
    ax.set_yticklabels([str(m)[:32] for m in top["bus_manufacturer"]], fontsize=8)
    ax.xaxis.set_major_formatter(_m); ax.set_xlabel("$m gross (all orbits)")
    ax.set_title("Top bus manufacturers by exposure")
    ax.grid(axis="x", color=RULE, linewidth=0.6, zorder=0)
    return _save(fig, Path(outdir) / "concentration.png")


def generate_all(grid, sw, outdir, changes=None, closing_total=None):
    outdir = Path(outdir); outdir.mkdir(parents=True, exist_ok=True)
    out = {}
    for key, fn in [
        ("exec_ranking", lambda: exec_ranking(grid, outdir)),
        ("scenario_compare", lambda: scenario_compare(grid, outdir)),
        ("waterfall", lambda: waterfall(grid, outdir)),
        ("concentration", lambda: concentration(sw, outdir)),
    ]:
        try:
            p = fn()
            if p:
                out[key] = p
        except Exception:
            pass
    if changes is not None and closing_total is not None:
        try:
            p = exposure_bridge(changes, closing_total, outdir)
            if p:
                out["exposure_bridge"] = p
        except Exception:
            pass
    return out
