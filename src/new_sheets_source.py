# New sheet builders to append to excel_report.py — modeled on PR RDS house standard

def _methodology(wb, params, changes):
    """Waterfall logic per entity + change log. Mirrors PR RDS Methodology tab."""
    ws = wb.create_sheet("Methodology")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 110
    ws["B2"] = "Space RDS  —  Methodology"; ws["B2"].font = F_TITLE
    ws["B3"] = (f"As at {params.as_at}   |   {params.quarter}   |   "
                "Scenario definitions, waterfall logic and change log.")
    ws["B3"].font = F_SUB

    r = _section(ws, 5, "Scenario definitions")
    scen = [
        ("Proton Flare", "5% insured loss applied to all on-risk GEO-GSO spacecraft. "
         "No time decay. Single solar-event scenario."),
        ("Space Weather", "Worst single bus-manufacturer total across ALL on-risk orbits "
         "(raw gross per-S/C, before equity/netting). Captures correlated manufacturer "
         "design-defect / solar exposure."),
        ("Generic Defect", "50% insured loss on GEO-GSO + MEO spacecraft, time-decayed by "
         "Risk Period Factor (RPF). Common-cause manufacturing defect."),
        ("Space Debris", "Orbit-specific damage ratios (LEO 40%, MEO 10%, GEO 5%) applied "
         "to on-risk spacecraft in each orbit band. Kessler-type collision scenario."),
        ("Max Risk", "Single largest on-risk spacecraft by gross per-S/C exposure. "
         "Worst single-risk loss."),
    ]
    r = _hdr(ws, r, 2, ["Scenario", "Definition"], [16, 108])
    for k, (s, d) in enumerate(scen):
        _cell(ws, r, 2, s, alt=k % 2 == 0)
        c = _cell(ws, r, 3, d, alt=k % 2 == 0)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 42
        r += 1
    r += 1

    r = _section(ws, r, "Netting waterfall logic")
    wf = [
        ("FUL", "Gross → External QS (outwards RI slots by inception) → IGR QS 50% ceded "
         "to FIBL → IGR XoL $245m xs $50m → Net retained."),
        ("FIID", "Same cascade as FUL with IGR QS 85% and IGR XoL $72.5m xs $7.5m."),
        ("FIHL", "Group-consolidated gross — external cessions only (no intra-group IGR). "
         "Equity-share basis."),
        ("S3123", "Syndicate 3123 inwards 20% QS on consortium-eligible layers within the "
         "treaty window; IG equity slice applied. Excluded spacecraft per parameters."),
    ]
    r = _hdr(ws, r, 2, ["Entity", "Cascade"], [16, 108])
    for k, (e, d) in enumerate(wf):
        _cell(ws, r, 2, e, alt=k % 2 == 0)
        c = _cell(ws, r, 3, d, alt=k % 2 == 0)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30
        r += 1
    r += 1

    r = _section(ws, r, "Population & exclusion logic")
    logic = [
        "On-risk filter: Inception ≤ as-at ≤ Off-risk date (re-applied after any data correction).",
        "Consortium placing-basis layers excluded in-pipeline (SQL view filters only "
        "'Consortium Declaration'; pipeline also drops plain 'Consortium').",
        "NULL-orbit / no-Seradata-ID layers excluded from all scenario calculations "
        "(revised exposure set to 0); retained on Python Adjustments tab for audit.",
        "Manual add/remove layers sourced from rds.manually_controlled_rds_layers (SQL).",
        "Data corrections (e.g. wrong coverage period) override the source field and re-filter.",
    ]
    for k, t in enumerate(logic):
        c = _cell(ws, r, 2, "•  " + t, alt=k % 2 == 0)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 28
        r += 1
    r += 1

    r = _section(ws, r, "Change log")
    r = _hdr(ws, r, 2, ["Quarter", "Change"], [16, 108])
    if changes:
        s = changes["layers"]["summary"]
        log = f"vs prior run ({changes['prior_as_at']}): {s['new_layers']} new, " \
              f"{s['dropped_layers']} dropped layers."
    else:
        log = "Initial automation deployment — no prior quarter to compare."
    _cell(ws, r, 2, params.quarter)
    c = _cell(ws, r, 3, log); c.alignment = Alignment(wrap_text=True)
    r += 2
    ws.cell(row=r, column=2,
            value="Pipeline: run_space_rds.py → ingest · engine · scenarios · netting · "
                  "outputs. Source of truth: SQL view rds.vw_SpaceRDS_OnRisk.").font = F_SUB


def _audit(wb, per_layer, params, source, recon, excluded):
    """Extract row counts, SQL refs, validation findings. Mirrors PR RDS Audit tab."""
    ws = wb.create_sheet("Audit")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 80
    ws["B2"] = "SPACE RDS  —  Audit"; ws["B2"].font = F_TITLE
    ws["B3"] = (f"As at {params.as_at}   |   Run {dt.datetime.now():%Y-%m-%d %H:%M:%S}"
                f"   |   source: {source}   |   SQL refs, row counts, validation.")
    ws["B3"].font = F_SUB

    r = _section(ws, 5, "Extract — connections & row counts")
    r = _hdr(ws, r, 2, ["Dataset", "Rows", "SQL source"], [40, 12, 80])
    n_layers = len(per_layer)
    n_excl = len(excluded) if excluded is not None else 0
    rows = [
        ("On-risk Space layers (final population)", n_layers,
         "rds.vw_SpaceRDS_OnRisk (LON-SQLP-V005 / SpaceTrax_Data), as-at filtered"),
        ("Excluded / adjusted layers", n_excl,
         "Pipeline exclusions: consortium, NULL-orbit, data corrections"),
        ("IGR cession rates", "—",
         "param_ig_cessions_qs / param_ig_cessions_xol by entity·code·year"),
        ("External RI cessions", "—", "param_ri_cessions by inception window"),
        ("Manual add/remove layers", "—", "rds.manually_controlled_rds_layers (Action col)"),
    ]
    for k, (d, n, s) in enumerate(rows):
        alt = k % 2 == 0
        _cell(ws, r, 2, d, alt=alt)
        c = _cell(ws, r, 3, n, alt=alt); c.alignment = Alignment(horizontal="right")
        _cell(ws, r, 4, s, alt=alt); r += 1
    r += 1

    r = _section(ws, r, "Validation findings")
    findings = []
    if recon is not None and "status" in getattr(recon, "columns", []):
        n_bad = int((recon["status"] == "MISMATCH").sum())
        n_ok = int((recon["status"] == "OK").sum())
        findings.append(f"Reconciliation vs workbook: {n_ok} checks OK, {n_bad} mismatches."
                        if n_bad else f"Reconciliation vs workbook: all {n_ok} checks OK.")
    null_orbit = int((per_layer["orbit"].astype(str).isin(["NULL", "None", "nan"])).sum())
    findings.append(f"NULL-orbit layers in final population: {null_orbit} (target 0 — "
                    "excluded upstream)." if null_orbit == 0
                    else f"NULL-orbit layers: {null_orbit} excluded, see Python Adjustments.")
    bad = int((per_layer["per_sc"] <= 0).sum())
    findings.append(f"Zero/negative exposure rows: {bad}." if bad
                    else "No zero/negative exposure rows.")
    if not findings:
        findings = ["No data-quality issues detected."]
    for k, t in enumerate(findings):
        c = _cell(ws, r, 2, t, alt=k % 2 == 0)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        r += 1
    r += 1

    r = _section(ws, r, "Population reconciliation")
    total = per_layer["per_sc"].sum()
    r = _kv(ws, r, "Final on-risk layers", n_layers)
    r = _kv(ws, r, "Total signed exposure (per-S/C)", total, money=True)
    r = _kv(ws, r, "Layers adjusted out / in (Python Adjustments)", n_excl)
    r += 1
    ws.cell(row=r, column=2,
            value="Full per-layer extract on the Per Layer tab. All summary tabs are live "
                  "Excel formulas referencing it (clickable audit trail).").font = F_SUB


def _python_adjustments(wb, per_layer, excluded):
    """Merged Excluded + Manual Overrides — all pipeline add/remove with reasons."""
    ws = wb.create_sheet("Python Adjustments")
    ws.sheet_view.showGridLines = False
    _title(ws, "Python Adjustments",
           "automated layer removals & additions, with documented reasons · for audit")

    # Removals (from the engine's excluded set)
    rem = excluded.copy() if (excluded is not None and len(excluded)) else None
    adds = per_layer[per_layer.get("override_action", pd.Series(dtype=str))
                     .astype(str).str.strip().str.lower() == "add layer"] \
        if "override_action" in per_layer.columns else per_layer.iloc[0:0]

    r = _section(ws, 5, "Summary")
    n_rem = len(rem) if rem is not None else 0
    rem_exp = float(rem["per_sc"].fillna(0).sum()) if rem is not None else 0.0
    add_exp = float(adds["per_sc"].fillna(0).sum()) if len(adds) else 0.0
    r = _hdr(ws, r, 2, ["Adjustment", "Layers", "Exposure impact"], [40, 10, 18])
    _cell(ws, r, 2, "Removed (excluded from calculations)")
    _cell(ws, r, 3, n_rem); c = _cell(ws, r, 4, -rem_exp, money=True); r += 1
    _cell(ws, r, 2, "Added (manual inclusions from SQL table)", alt=True)
    _cell(ws, r, 3, len(adds), alt=True)
    _cell(ws, r, 4, add_exp, alt=True, money=True); r += 1
    _cell(ws, r, 2, "Net adjustment", bold=True)
    _cell(ws, r, 3, len(adds) - n_rem, bold=True)
    _cell(ws, r, 4, add_exp - rem_exp, money=True, bold=True); r += 2

    # Removals detail
    if rem is not None and len(rem):
        r = _section(ws, r, f"Removals ({n_rem}) — excluded from all scenario calculations")
        cols = [("program_id", "Program", 11), ("layer_id", "Layer", 8),
                ("entity", "Entity", 9), ("spacecraft_name", "Spacecraft", 24),
                ("orbit", "Orbit", 10), ("per_sc", "Orig Exposure", 15),
                ("excluded_reason", "Reason for removal", 46)]
        r = _hdr(ws, r, 2, [h for _, h, _ in cols], [w for _, _, w in cols])
        rem2 = rem.sort_values("excluded_reason") if "excluded_reason" in rem.columns else rem
        for k, (_, row) in enumerate(rem2.iterrows()):
            alt = k % 2 == 0
            for j, (field, _, _) in enumerate(cols):
                v = row.get(field)
                if field == "spacecraft_name" and (v is None or str(v) == "nan"):
                    v = "(none)"
                if field == "excluded_reason" and (v is None or str(v) == "nan" or str(v) == "None"):
                    v = "Out of scenario scope"
                _cell(ws, r, 2 + j, v, alt=alt, money=(field == "per_sc"))
            r += 1
        r += 1

    # Additions detail
    if len(adds):
        r = _section(ws, r, f"Additions ({len(adds)}) — manual inclusions applied via SQL")
        cols = [("program_id", "Program", 11), ("layer_id", "Layer", 8),
                ("entity", "Entity", 9), ("spacecraft_name", "Spacecraft", 24),
                ("orbit", "Orbit", 10), ("inception", "Inception", 13),
                ("per_sc", "Exposure", 15)]
        r = _hdr(ws, r, 2, [h for _, h, _ in cols], [w for _, _, w in cols])
        for k, (_, row) in enumerate(adds.sort_values("per_sc", ascending=False).iterrows()):
            alt = k % 2 == 0
            for j, (field, _, _) in enumerate(cols):
                v = row.get(field)
                if isinstance(v, (dt.date, dt.datetime)):
                    v = str(v)
                if field == "spacecraft_name" and (v is None or str(v) == "nan"):
                    v = "(none)"
                _cell(ws, r, 2 + j, v, alt=alt, money=(field == "per_sc"))
            r += 1
        r += 1

    ws.cell(row=r, column=2,
            value="Reasons: 'No orbit' = satellite not yet launched / no Seradata ID; "
                  "'Consortium placing basis' = consortium-level aggregate (excluded per "
                  "methodology); 'Off-risk after data correction' = coverage period "
                  "corrected to true expiry. Additions sourced from "
                  "rds.manually_controlled_rds_layers.").font = F_SUB
