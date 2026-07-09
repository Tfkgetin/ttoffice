-- ===========================================================================
-- Diagnostic: why does Arabsat successor 382904 NOT flow through the RDS extract
-- even though Layer_Details_v shows it IsBound=1 (~$22.4m)?
--
-- The RDS extract keeps a layer only if it clears TWO gates that Layer_Details_v
-- alone does not impose (see sql/space_rds_onrisk_asat.sql):
--   1. Controlling_Body must be IG / MGU / Consortium (Prequel.Lookups)
--   2. rds.param_consortium_splits must have an MGU row covering the layer's
--      inception  (cs.[Start Date] <= inception AND cs.[End Date] >= inception)
-- A Direct Open Market layer with no such controlling body / split row is
-- silently dropped — the Turksat-class gap. Run each block and read the notes.
-- Compare against the EXPIRING programme 344664, which DOES flow, to see which
-- gate differs.
-- ===========================================================================

-- Q1 — Controlling body per layer (the gate at join #1).
--      If 382904's rows show a Controlling_Body NOT in (IG, MGU, Consortium),
--      that is why they never reach BaseData.
SELECT  P.ProgramId,
        l.Controlling_BodyID,
        b.Controlling_Body,
        P.Placing_Basis,
        COUNT(*)                               AS layers,
        SUM(CAST(P.IsBound AS INT))            AS bound_layers,
        MIN(P.Inception)                       AS earliest_inc,
        MAX(P.Expiry)                          AS latest_exp
FROM        [Prequel_Reporting].[Prequel].[Layer_Details_v] P
JOIN        Prequel.Data.Layers_t l              ON l.LayerID = P.LayerID
LEFT JOIN   Prequel.Lookups.Controlling_Bodies_t b ON b.Controlling_BodyID = l.Controlling_BodyID
WHERE   P.ProgramId IN (344664, 382904)          -- expiring vs successor
GROUP BY P.ProgramId, l.Controlling_BodyID, b.Controlling_Body, P.Placing_Basis
ORDER BY P.ProgramId, b.Controlling_Body;

-- Q2 — Consortium-split coverage (the gate at join #2), MGU controlling body.
--      For each 382904 layer, is there an MGU param_consortium_splits row whose
--      [Start Date]..[End Date] spans its inception? NULL cession = NO row = the
--      layer is dropped by the INNER JOIN.
SELECT  P.ProgramId, P.LayerId, P.Inception,
        cs.[Start Date] AS split_start,
        cs.[End Date]   AS split_end,
        cs.[Cession]    AS mgu_cession,
        CASE WHEN cs.[Cession] IS NULL
             THEN 'DROPPED — no MGU split row covers this inception'
             ELSE 'ok — split row present' END AS verdict
FROM        [Prequel_Reporting].[Prequel].[Layer_Details_v] P
LEFT JOIN   [SpaceTrax_Data].[rds].param_consortium_splits cs
        ON  cs.[Start Date] <= P.Inception
        AND cs.[End Date]   >= P.Inception
        AND cs.[Controlling Body] = 'MGU'
WHERE   P.ProgramId = 382904
ORDER BY P.LayerId;

-- Q3 — What MGU split rows exist right now, and do any cover 2026-07-01?
--      (If the latest row ends 2026-06-30, a 2026-07-01 renewal is uncovered —
--      exactly the gap that hid Turksat until a new row was added.)
SELECT  [Controlling Body], [Start Date], [End Date], [Cession]
FROM    [SpaceTrax_Data].[rds].param_consortium_splits
WHERE   [Controlling Body] = 'MGU'
ORDER BY [Start Date];

-- Q4 — Mapping-code / program-type gate (a lesser gate).
--      The extract keeps only Mapping_Code IN ('ASO','ASC') and
--      Program_Type NOT IN ('LVFO','Satellite Launch'). Confirm 382904 qualifies.
SELECT  P.ProgramId, P.Mapping_Code, P.Program_Type, COUNT(*) AS layers
FROM    [Prequel_Reporting].[Prequel].[Layer_Details_v] P
WHERE   P.ProgramId IN (344664, 382904)
GROUP BY P.ProgramId, P.Mapping_Code, P.Program_Type
ORDER BY P.ProgramId;

-- ---------------------------------------------------------------------------
-- READING THE RESULTS
--   * Q1 shows Controlling_Body = 'Direct' (or blank) for 382904  → gate #1
--     drops it. Fix: it must be booked under IG/MGU/Consortium to flow, OR
--     carry it via renewal_rollforward (current stop-gap).
--   * Q2 shows mgu_cession = NULL for the 2026-07-01 layers                → gate #2
--     drops it. Fix: add an MGU param_consortium_splits row covering
--     2026-07-01.. (same fix that recovered Turksat).
--   * Q3 latest [End Date] < 2026-07-01                                    → confirms the
--     split table doesn't reach the new annual tranche yet.
--   * Q4 anything other than ASO/ASC or an excluded Program_Type            → a third
--     reason it would drop.
-- Whichever gate fires is the real cause; fix that and 382904 flows at its
-- true bound value (~$22.4m), the rollforward safeguard auto-retires, and
-- Arabsat leaves 'Renewals in progress' on its own.
-- ---------------------------------------------------------------------------
