-- Coverage audit — EVERY bound, in-scope space layer that SHOULD be considered
-- for the RDS, with a flag for whether the extract's MGU consortium-split gate
-- would silently DROP it.
--
-- The on-risk extract INNER-JOINs rds.param_consortium_splits (MGU) on each
-- layer's inception, so any consortium/IG layer whose inception is not covered
-- by an MGU split row is silently dropped (this is what hid Turksat 6A and
-- Hellas Sat 382815). This query mirrors the extract's METHODOLOGY filters
-- (ASO/ASC, not LVFO/Satellite Launch, not Consortium Declaration, on-risk at
-- the as-at) but LEFT-joins the split table so split-missing layers are RETAINED
-- and flagged (has_mgu_split = 0) instead of dropped.
--
-- The pipeline (ingest.coverage_audit) reads this, and flags any row with
-- has_mgu_split = 0 whose programme is NOT represented in the final book and NOT
-- carried by a config rule — i.e. a potential MISS. {as_at} is injected.
DECLARE @AsAt DATE = '{as_at}';

SELECT
    P.ProgramId,
    P.LayerId,
    b.Controlling_Body,
    sc.[Orbit Category]                                   AS Orbit,
    ls.[SatelliteNameAlternative]                         AS SpacecraftId,
    sc.[Spacecraft Name]                                  AS SpacecraftName,
    ISNULL(P.[Client_Name],'') + ' ' + ISNULL(P.[Program_Name],'') AS Programme,
    P.Inception,
    P.Expiry,
    P.Mapping_Code,
    P.Signed_Exposure_USD,
    CASE WHEN cs.[Controlling Body] IS NULL THEN 0 ELSE 1 END AS has_mgu_split
FROM        [Prequel_Reporting].[Prequel].[Layer_Details_v] P
JOIN        [Prequel].[Data].[Layers_t]                 l  ON l.LayerID = P.LayerID
JOIN        [Prequel].[Lookups].[Controlling_Bodies_t]  b  ON b.Controlling_BodyID = l.Controlling_BodyID
LEFT JOIN   [Prequel].[Data].[Space_Layers_t]           ls ON ls.LayerID = P.LayerID
LEFT JOIN   [SpaceTrax_Data].[dbo].[tbl_SpaceCraft]     sc ON sc.[Seradata Spacecraft ID] = ls.[SatelliteNameAlternative]
LEFT JOIN   [SpaceTrax_Data].[rds].[param_consortium_splits] cs
       ON cs.[Start Date] <= l.Inception AND cs.[End Date] >= l.Inception
      AND cs.[Controlling Body] = 'MGU'
WHERE P.IsBound = 1
  AND b.Controlling_Body IN ('Consortium','IG','MGU')
  AND P.Inception <= @AsAt
  AND P.Expiry    >= @AsAt
  AND P.Mapping_Code   IN ('ASO','ASC')
  AND P.Program_Type   NOT IN ('LVFO','Satellite Launch')
  AND P.Placing_Basis  NOT IN ('Consortium Declaration')
ORDER BY has_mgu_split, b.Controlling_Body, P.Inception;
