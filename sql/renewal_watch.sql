-- Renewal watch: space layers with their forward renewal pointer, from the
-- latest PBI layer snapshot. Used by the renewal-policy check to classify the
-- quarter's expiring / dropped layers (renewed-and-bound / in-progress / NTU /
-- no-pointer) independent of the on-risk extract. Keyed on Program Id + Layer Id
-- (= the pipeline's layer_key). Sentinel 0 → NULL. Read-only, no date param —
-- the pipeline matches rows to its own dropped set by layer_key.
SELECT
    [Program Id]                         AS ProgramId,
    [Layer Id]                           AS LayerId,
    [Inception]                          AS Inception,
    [Expiry]                             AS Expiry,
    [UW Status]                          AS UW_Status,
    NULLIF([Renewed To Program Id],  0)  AS Renewed_To_Program_Id,
    NULLIF([Renewed From Program Id], 0) AS Renewed_From_Program_Id,
    [Renewed To UW Status]               AS Renewed_To_UW_Status,
    [Renewed To Inception]               AS Renewed_To_Inception
FROM [J].[Pbi].[Layers_t]
WHERE [Snapshot Id] = (SELECT MAX([Snapshot Id]) FROM [J].[Pbi].[Layers_t])
  AND [Minor Class] = 'Space';
