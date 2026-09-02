-- ===========================================================================
-- Diagnostic: why is a given LayerId absent from the RDS population?
--
-- A layer can be missing for two different reasons, and only ONE of them is
-- visible in the workbook:
--   * the PIPELINE dropped it (placing basis / on-risk re-filter / null orbit /
--     re-added duplicate)  -> it appears on the Excluded tab with a reason;
--   * the EXTRACT never selected it                                 -> it
--     appears NOWHERE, because there was nothing to log.
-- This script diagnoses the second case. It replays every gate in
-- sql/space_rds_onrisk_asat.sql against a list of layer ids and names the gate
-- that dropped each one.
--
-- Set @AsAt to the run's as-at date and @Layers to the ids you are chasing.
-- Q0 is the answer; Q1-Q5 explain the gates Q0 flags.
-- ===========================================================================

DECLARE @AsAt DATE = '2026-07-01';

IF OBJECT_ID('tempdb..#Layers') IS NOT NULL DROP TABLE #Layers;
CREATE TABLE #Layers (LayerId INT PRIMARY KEY);
INSERT INTO #Layers (LayerId) VALUES
    (475187), (375916), (375920), (359480), (359481), (341461), (340163);

-- ---------------------------------------------------------------------------
-- Q0 — MASTER VERDICT. One row per layer; `verdict` lists every gate it fails.
--      An empty verdict means the extract SHOULD have kept it, in which case
--      the layer is being lost downstream (see the note at the foot of Q0).
-- ---------------------------------------------------------------------------
WITH Base AS (
    SELECT  P.ProgramId, P.LayerId, P.Mapping_Code, P.Program_Type,
            P.Placing_Basis, P.IsBound,
            P.Inception  AS ldv_inception,
            P.Expiry     AS ldv_expiry,
            P.Signed_Exposure_USD,
            l.Inception  AS lt_inception,     -- NOTE: the split join uses THIS
            l.Expiry     AS lt_expiry,
            l.Controlling_BodyID,
            b.Controlling_Body,
            ls.SatelliteNameAlternative,
            ls.Coverage_Period_Months,
            c.Description                     AS cover_desc,
            sc.[Launch Date]                  AS launch_date,
            sc.[Orbit Category]               AS orbit,
            CASE WHEN l.LayerID IS NULL THEN 0 ELSE 1 END AS in_layers_t
    FROM        [Prequel_Reporting].[Prequel].[Layer_Details_v] P
    INNER JOIN  #Layers T                             ON T.LayerId = P.LayerId
    LEFT JOIN   Prequel.Data.Layers_t l               ON l.LayerID = P.LayerID
    LEFT JOIN   Prequel.Lookups.Controlling_Bodies_t b ON b.Controlling_BodyID = l.Controlling_BodyID
    LEFT JOIN   Prequel.Data.Space_Layers_t ls        ON ls.LayerID = P.LayerID
    LEFT JOIN   Prequel.Lookups.Coverage_Types_t c    ON ls.CoverRequired = c.[Coverage_TypeID]
    LEFT JOIN   SpaceTrax_Data.dbo.tbl_SpaceCraft sc  ON sc.[Seradata Spacecraft ID] = ls.[SatelliteNameAlternative]
),
-- Reproduce SpacecraftData's own WHERE clause: a spacecraft only supplies
-- Coverage / MonthsOnRisk when it has a launch date AND a numeric id.
Derived AS (
    SELECT B.*,
           CASE WHEN B.launch_date IS NOT NULL
                 AND ISNUMERIC(B.SatelliteNameAlternative) = 1
                THEN ISNULL(B.cover_desc, 'In-Orbit') END          AS Coverage,
           CASE WHEN B.launch_date IS NOT NULL
                 AND ISNUMERIC(B.SatelliteNameAlternative) = 1
                THEN CASE
                      WHEN B.Coverage_Period_Months IS NOT NULL THEN B.Coverage_Period_Months
                      WHEN B.cover_desc = 'L + 12 Months' THEN 12
                      WHEN B.cover_desc = 'L + 24 Months' THEN 24
                      WHEN B.cover_desc = 'L + 6 Months'  THEN 6
                      ELSE CAST(ROUND(CAST(DATEDIFF(DAY, B.lt_inception, B.lt_expiry)
                                           AS DECIMAL(9,2)) / 365 * 12, 0) AS INT)
                     END END                                       AS MonthsOnRisk
    FROM Base B
),
Dates AS (
    SELECT D.*,
           CAST(CASE WHEN D.Coverage <> 'In-Orbit' THEN D.launch_date
                     ELSE D.ldv_inception END AS DATE)             AS On_Risk_Date,
           CAST(CASE WHEN D.Coverage <> 'In-Orbit'
                     THEN DATEADD(MONTH, D.MonthsOnRisk, D.launch_date)
                     ELSE D.ldv_expiry END AS DATE)                AS Off_Risk_Date
    FROM Derived D
),
Gates AS (
    SELECT X.*,
           -- G1  row present in Layers_t (INNER JOIN in both branches)
           X.in_layers_t                                              AS g1_layers_t,
           -- G2  controlling body must be Consortium (branch A) or IG/MGU (branch B)
           CASE WHEN X.Controlling_Body IN ('Consortium','IG','MGU')
                THEN 1 ELSE 0 END                                     AS g2_ctrl_body,
           -- G3  INNER JOIN rds.param_mapping_code
           CASE WHEN EXISTS (SELECT 1 FROM SpaceTrax_Data.rds.param_mapping_code M
                             WHERE M.Mapping_Code = X.Mapping_Code)
                THEN 1 ELSE 0 END                                     AS g3_mapping_tbl,
           -- G4  INNER JOIN param_consortium_splits on Layers_t.Inception (MGU)
           CASE WHEN EXISTS (SELECT 1 FROM SpaceTrax_Data.rds.param_consortium_splits S
                             WHERE S.[Controlling Body] = 'MGU'
                               AND S.[Start Date] <= X.lt_inception
                               AND S.[End Date]   >= X.lt_inception)
                THEN 1 ELSE 0 END                                     AS g4_mgu_split,
           CASE WHEN X.IsBound = 1 THEN 1 ELSE 0 END                  AS g5_isbound,
           CASE WHEN X.Off_Risk_Date >= @AsAt THEN 1 ELSE 0 END       AS g6_offrisk,
           CASE WHEN X.ldv_inception <= @AsAt THEN 1 ELSE 0 END       AS g7_inception,
           -- G8-G10: NOT IN / IN against a NULL column yields UNKNOWN, which is
           -- NOT true — so a NULL here silently drops the layer.
           CASE WHEN X.Mapping_Code IN ('ASO','ASC') THEN 1 ELSE 0 END AS g8_mapping_code,
           CASE WHEN X.Program_Type NOT IN ('LVFO','Satellite Launch')
                THEN 1 ELSE 0 END                                      AS g9_prog_type,
           CASE WHEN X.Placing_Basis NOT IN ('Consortium Declaration')
                THEN 1 ELSE 0 END                                      AS g10_placing,
           -- G11  manual-control table. NOTE the branches differ: the Consortium
           --      branch excludes only Action IN ('Add Layer','Remove Layer');
           --      the IG branch excludes ANY row in the table, whatever Action.
           CASE WHEN EXISTS (SELECT 1 FROM SpaceTrax_Data.rds.manually_controlled_rds_layers MC
                             WHERE MC.LayerID = X.LayerId)
                THEN 0 ELSE 1 END                                      AS g11_manual_any,
           CASE WHEN EXISTS (SELECT 1 FROM SpaceTrax_Data.rds.manually_controlled_rds_layers MC
                             WHERE MC.LayerID = X.LayerId
                               AND MC.[Action] IN ('Add Layer','Remove Layer'))
                THEN 0 ELSE 1 END                                      AS g11_manual_addrem,
           -- G12  the outer BaseData window, on the DERIVED on-risk date
           CASE WHEN @AsAt >= X.On_Risk_Date THEN 1 ELSE 0 END         AS g12_onrisk
    FROM Dates X
)
SELECT  ProgramId, LayerId, Controlling_Body, Mapping_Code, Program_Type,
        Placing_Basis, IsBound, Coverage, orbit,
        ldv_inception, lt_inception, On_Risk_Date, Off_Risk_Date,
        Signed_Exposure_USD,
        LTRIM(
          CASE WHEN g1_layers_t     = 0 THEN ' | no Layers_t row'                      ELSE '' END +
          CASE WHEN g2_ctrl_body    = 0 THEN ' | controlling body not IG/MGU/Consortium (=' + ISNULL(Controlling_Body,'NULL') + ')' ELSE '' END +
          CASE WHEN g3_mapping_tbl  = 0 THEN ' | mapping code not in param_mapping_code' ELSE '' END +
          CASE WHEN g4_mgu_split    = 0 THEN ' | NO MGU split row covers Layers_t.Inception' ELSE '' END +
          CASE WHEN g5_isbound      = 0 THEN ' | IsBound <> 1'                          ELSE '' END +
          CASE WHEN g6_offrisk      = 0 THEN ' | off-risk before as-at'                 ELSE '' END +
          CASE WHEN g7_inception    = 0 THEN ' | inception after as-at'                 ELSE '' END +
          CASE WHEN g8_mapping_code = 0 THEN ' | Mapping_Code not ASO/ASC (=' + ISNULL(Mapping_Code,'NULL') + ')' ELSE '' END +
          CASE WHEN g9_prog_type    = 0 THEN ' | Program_Type LVFO/Launch or NULL (=' + ISNULL(Program_Type,'NULL') + ')' ELSE '' END +
          CASE WHEN g10_placing     = 0 THEN ' | Placing_Basis Consortium Declaration or NULL (=' + ISNULL(Placing_Basis,'NULL') + ')' ELSE '' END +
          CASE WHEN g11_manual_any  = 0 THEN ' | present in manually_controlled_rds_layers' ELSE '' END +
          CASE WHEN g12_onrisk      = 0 THEN ' | on-risk date after as-at'              ELSE '' END,
        ' |') AS verdict
FROM Gates
ORDER BY ProgramId, LayerId;
-- If `verdict` is EMPTY the extract should have kept the layer. Then either
-- (a) it IS in the extract and the pipeline dropped it later — check the
--     Excluded tab and the run's null_orbit_layers.csv; or
-- (b) the layer is duplicated / superseded — see Q5.

-- ---------------------------------------------------------------------------
-- Q1 — THE NULL TRAP. This is the gate the older diagnostic does not test.
--      `Placing_Basis NOT IN ('Consortium Declaration')` evaluates to UNKNOWN
--      when Placing_Basis IS NULL, and UNKNOWN is not TRUE, so the row is
--      dropped exactly as if it were a consortium declaration. The same applies
--      to Program_Type. A layer can therefore be filtered out by a field that
--      is simply blank.
-- ---------------------------------------------------------------------------
SELECT  P.ProgramId, P.LayerId,
        P.Placing_Basis, P.Program_Type, P.Mapping_Code,
        CASE WHEN P.Placing_Basis IS NULL THEN 'DROPPED — Placing_Basis is NULL' END AS placing_verdict,
        CASE WHEN P.Program_Type  IS NULL THEN 'DROPPED — Program_Type is NULL'  END AS progtype_verdict,
        CASE WHEN P.Mapping_Code  IS NULL THEN 'DROPPED — Mapping_Code is NULL'  END AS mapping_verdict
FROM    [Prequel_Reporting].[Prequel].[Layer_Details_v] P
JOIN    #Layers T ON T.LayerId = P.LayerId
ORDER BY P.ProgramId, P.LayerId;

-- ---------------------------------------------------------------------------
-- Q2 — MGU split coverage, joined on the RIGHT date.
--      The extract joins param_consortium_splits on Layers_t.Inception
--      (cs.[Start Date] <= l.inception), NOT on Layer_Details_v.Inception.
--      Where the two differ, testing the wrong one gives a false "ok".
-- ---------------------------------------------------------------------------
SELECT  P.ProgramId, P.LayerId,
        P.Inception  AS ldv_inception,
        l.Inception  AS layers_t_inception,
        CASE WHEN P.Inception <> l.Inception THEN '<-- DATES DIFFER' END AS date_mismatch,
        S.[Start Date], S.[End Date], S.[Cession],
        CASE WHEN S.[Cession] IS NULL
             THEN 'DROPPED — no MGU split row covers Layers_t.Inception'
             ELSE 'ok' END AS verdict
FROM        [Prequel_Reporting].[Prequel].[Layer_Details_v] P
JOIN        #Layers T ON T.LayerId = P.LayerId
LEFT JOIN   Prequel.Data.Layers_t l ON l.LayerID = P.LayerID
LEFT JOIN   SpaceTrax_Data.rds.param_consortium_splits S
        ON  S.[Controlling Body] = 'MGU'
        AND S.[Start Date] <= l.Inception
        AND S.[End Date]   >= l.Inception
ORDER BY P.ProgramId, P.LayerId;

-- Q2b — every MGU split period, to see where the gaps are.
SELECT [Controlling Body], [Start Date], [End Date], [Cession]
FROM   SpaceTrax_Data.rds.param_consortium_splits
WHERE  [Controlling Body] = 'MGU'
ORDER BY [Start Date];

-- ---------------------------------------------------------------------------
-- Q3 — THE BRANCH ASYMMETRY on the manual-control table.
--      Consortium branch:  excluded only if Action IN ('Add Layer','Remove Layer')
--      IG / MGU branch:    excluded if the layer is in the table AT ALL
--      So an IG layer sitting in that table under any other Action (or a NULL
--      Action) is dropped, with nothing on the Excluded tab to say so.
-- ---------------------------------------------------------------------------
SELECT  MC.ProgramId, MC.LayerID, MC.[Action],
        MC.Consortium, MC.IGOnly, MC.LLoydsOnly,
        b.Controlling_Body,
        CASE WHEN b.Controlling_Body IN ('IG','MGU')
             THEN 'DROPPED from the IG branch (any Action excludes)'
             WHEN MC.[Action] IN ('Add Layer','Remove Layer')
             THEN 'DROPPED from the Consortium branch (Add/Remove)'
             ELSE 'kept' END AS verdict
FROM        SpaceTrax_Data.rds.manually_controlled_rds_layers MC
JOIN        #Layers T ON T.LayerId = MC.LayerID
LEFT JOIN   Prequel.Data.Layers_t l ON l.LayerID = MC.LayerID
LEFT JOIN   Prequel.Lookups.Controlling_Bodies_t b ON b.Controlling_BodyID = l.Controlling_BodyID;

-- Q3b — a manual ADD only survives this filter:
--       ((Consortium = 1 AND LLoydsOnly <> 1) OR IGOnly = 1)
--       A NULL LLoydsOnly makes `LLoydsOnly <> 1` UNKNOWN, so a Consortium add
--       with a NULL LLoydsOnly is silently NOT added.
SELECT  MC.ProgramId, MC.LayerID, MC.[Action],
        MC.Consortium, MC.IGOnly, MC.LLoydsOnly,
        CASE WHEN ((MC.Consortium = 1 AND MC.LLoydsOnly <> 1) OR MC.IGOnly = 1)
             THEN 'passes' ELSE 'DROPPED — fails the ManualOverrides filter' END AS verdict
FROM    SpaceTrax_Data.rds.manually_controlled_rds_layers MC
JOIN    #Layers T ON T.LayerId = MC.LayerID;

-- ---------------------------------------------------------------------------
-- Q4 — The derived dates, shown alongside their inputs, so a NULL Off_Risk_Date
--      is visible. For launch cover the off-risk date is
--      DATEADD(MONTH, MonthsOnRisk, [Launch Date]); if MonthsOnRisk resolves to
--      NULL the result is NULL, `NULL >= @AsAt` is UNKNOWN, and the layer goes.
-- ---------------------------------------------------------------------------
SELECT  P.ProgramId, P.LayerId,
        c.Description                   AS cover_desc,
        ls.Coverage_Period_Months,
        ls.SatelliteNameAlternative,
        ISNUMERIC(ls.SatelliteNameAlternative) AS id_is_numeric,
        sc.[Launch Date],
        sc.[Orbit Category],
        P.Inception, P.Expiry, l.Inception AS lt_inception, l.Expiry AS lt_expiry,
        CASE WHEN sc.[Launch Date] IS NULL THEN 'no spacecraft row — falls back to Inception/Expiry'
             WHEN ISNUMERIC(ls.SatelliteNameAlternative) <> 1 THEN 'non-numeric Seradata id — falls back'
             WHEN ISNULL(c.Description,'In-Orbit') = 'In-Orbit' THEN 'in-orbit — uses Inception/Expiry'
             ELSE 'launch cover — uses Launch Date + MonthsOnRisk' END AS date_basis
FROM        [Prequel_Reporting].[Prequel].[Layer_Details_v] P
JOIN        #Layers T ON T.LayerId = P.LayerId
LEFT JOIN   Prequel.Data.Layers_t l            ON l.LayerID = P.LayerID
LEFT JOIN   Prequel.Data.Space_Layers_t ls     ON ls.LayerID = P.LayerID
LEFT JOIN   Prequel.Lookups.Coverage_Types_t c ON ls.CoverRequired = c.[Coverage_TypeID]
LEFT JOIN   SpaceTrax_Data.dbo.tbl_SpaceCraft sc ON sc.[Seradata Spacecraft ID] = ls.[SatelliteNameAlternative]
ORDER BY P.ProgramId, P.LayerId;

-- ---------------------------------------------------------------------------
-- Q5 — Sibling layers on the same programmes, so a layer that is genuinely
--      superseded (renewal, re-issued layer, split line) is obvious. Compare
--      the missing ids against the ones that DID flow.
-- ---------------------------------------------------------------------------
SELECT  P.ProgramId, P.LayerId,
        CASE WHEN T.LayerId IS NULL THEN '' ELSE '<-- CHASING' END AS chasing,
        P.Program_Name, P.Entity, P.Mapping_Code, P.Placing_Basis,
        b.Controlling_Body, P.IsBound,
        P.Inception, P.Expiry, P.Signed_Exposure_USD,
        ls.SatelliteNameAlternative AS spacecraft_id,
        sc.[Spacecraft Name], sc.[Orbit Category]
FROM        [Prequel_Reporting].[Prequel].[Layer_Details_v] P
LEFT JOIN   #Layers T ON T.LayerId = P.LayerId
LEFT JOIN   Prequel.Data.Layers_t l            ON l.LayerID = P.LayerID
LEFT JOIN   Prequel.Lookups.Controlling_Bodies_t b ON b.Controlling_BodyID = l.Controlling_BodyID
LEFT JOIN   Prequel.Data.Space_Layers_t ls     ON ls.LayerID = P.LayerID
LEFT JOIN   SpaceTrax_Data.dbo.tbl_SpaceCraft sc ON sc.[Seradata Spacecraft ID] = ls.[SatelliteNameAlternative]
WHERE   P.ProgramId IN (
            SELECT P2.ProgramId
            FROM   [Prequel_Reporting].[Prequel].[Layer_Details_v] P2
            JOIN   #Layers T2 ON T2.LayerId = P2.LayerId)
ORDER BY P.ProgramId, P.LayerId;
