-- Space RDS on-risk extract, as-at parameterised.
-- Body of [rds].[vw_SpaceRDS_OnRisk] with the DateParam CTE replaced by an
-- inline date ({as_at}, injected by the pipeline from the config's as_at_date).
-- Lets historical/interim runs use any as-at WITHOUT touching the shared
-- rds.param_as_at_date table.
-- CAVEATS for historical runs:
--   * manually_controlled_rds_layers is not dated: today's manual adds/removes
--     apply to any as-at you choose.
--   * Layer attributes (exposure, manufacturer, etc.) are as of today, not as
--     of the historical date - endorsements since then will show through.
--   * param tables must retain dated history (XoL terms join on the as-at).

WITH DateParam AS (
    SELECT CAST('{as_at}' AS DATE) AS As_At_Date
),
SpacecraftData AS (
     SELECT
      ls.[LayerID]
      ,[SatelliteNameAlternative] as SpacecraftId
      ,isnull(c.Description,'In-Orbit') as Coverage
      ,case when ls.Coverage_Period_Months IS NOT NULL THEN ls.Coverage_Period_Months
            when c.Description='L + 12 Months' then 12
            when c.Description='L + 24 Months' then 24
            when c.Description='L + 6 Months' then 6
            else cast(round(cast(DATEDIFF(DAY,l.Inception,l.Expiry) as decimal(9,2))/365*12,0) as int) end as MonthsOnRisk
      ,l.Expiry
      ,l.Inception
      ,sc.[Spacecraft Name] as [Spacecraft Name]
      ,sc.[Launch Date] as [Launch Date]
      ,sc.[Bus Type] as [Bus Type]
      ,sc.[Bus Family] as [Bus Family]
      ,sc.[Vehicle Family] as [Vehicle Family]
      ,sc.[Bus Manufacturer] as [Bus Manufacturer]
      ,sc.[Orbit Category] as [Orbit Category]
      ,sc.[Prime Manufacturer] as [Prime Manufacturer]
  FROM [Prequel].[Data].[Space_Layers_t] ls
  inner join [Prequel].[Data].Layers_t l on l.LayerID=ls.LayerID
  left JOIN [Prequel].[Lookups].[Coverage_Types_t] C ON ls.CoverRequired = C.[Coverage_TypeID]
  left join [SpaceTrax_Data].[dbo].[tbl_SpaceCraft] sc on sc.[Seradata Spacecraft ID]=ls.[SatelliteNameAlternative]
  where sc.[Launch Date] is not null and ISNUMERIC([SatelliteNameAlternative])=1
),
OriginalBaseDataConsortiumLayersAdjusted AS (
    SELECT
        P.[ProgramId], P.LayerId, Program_Type, P.Underwriting_Status,
        ISNULL(P.[Client_Name], '') + ' ' + ISNULL([Program_Name], '') AS [Program_Name],
        [Entity], P.[Mapping_Code], [Bus Family], [Bus Type], [Vehicle Family],
        Signed_Exposure_USD*cs_ig.Cession AS [Layer_Signed_Exposure_USD],
        Signed_Premium_USD*cs_ig.Cession as [Layer_Signed_Premium_USD],
        P.Inception, Coverage, P.Expiry, [Launch Date], [Orbit Category],
        MonthsOnRisk, IPO.SpacecraftId, [Spacecraft Name], [Bus Manufacturer], [Prime Manufacturer],
        P.Placing_Basis, b.Controlling_Body AS Controlling_Body,
        CAST(CASE WHEN Coverage <> 'In-Orbit' THEN [Launch Date] ELSE P.Inception END AS DATE) AS On_Risk_Date,
        CAST(CASE WHEN Coverage <> 'In-Orbit' THEN DATEADD(MONTH, [MonthsOnRisk], IPO.[Launch Date]) ELSE P.Expiry END AS DATE) AS Off_Risk_Date,
        NULL AS [Action], Null as Altitude_km, Null as Line_size_perc
    FROM [Prequel_Reporting].[Prequel].[Layer_Details_v] P
    left JOIN SpacecraftData IPO ON P.LayerId = IPO.LayerId
    JOIN [SpaceTrax_Data].[rds].[param_mapping_code] PMC ON P.Mapping_Code = PMC.Mapping_Code
    join Prequel.Data.Layers_t l on l.LayerID=P.LayerID
    join Prequel.Lookups.Controlling_Bodies_t b
        on b.Controlling_BodyID=l.Controlling_BodyID and b.Controlling_Body in ('Consortium')
    inner join [SpaceTrax_Data].[rds].param_consortium_splits cs_ig
        on cs_ig.[Start Date]<=l.inception and cs_ig.[End Date]>=l.inception
        AND cs_ig.[Controlling Body] = 'MGU'
    WHERE P.IsBound = 1
    AND (CASE WHEN Coverage <> 'In-Orbit' THEN DATEADD(MONTH, [MonthsOnRisk], IPO.[Launch Date]) ELSE P.Expiry END) >= (SELECT TOP 1 [As_At_Date] FROM DateParam)
    AND P.Inception <= (SELECT TOP 1 [As_At_Date] FROM DateParam)
    and p.Mapping_Code in ('ASO','ASC')
    AND p.Program_Type NOT IN ('LVFO','Satellite Launch')
    AND p.Placing_Basis NOT IN ('Consortium Declaration')
),
OriginalBaseDataIGLayersAdjusted AS (
    SELECT
        P.[ProgramId], P.LayerId, Program_Type, P.Underwriting_Status,
        ISNULL(P.[Client_Name], '') + ' ' + ISNULL([Program_Name], '') AS [Program_Name],
        [Entity], P.[Mapping_Code], [Bus Family], [Bus Type], [Vehicle Family],
        (isnull(Signed_Exposure_USD,0)) AS [Layer_Signed_Exposure_USD],
        (isnull(Signed_Premium_USD,0)) as [Layer_Signed_Premium_USD],
        P.Inception, Coverage, P.Expiry, [Launch Date], [Orbit Category],
        MonthsOnRisk, IPO.SpacecraftId, [Spacecraft Name], [Bus Manufacturer], [Prime Manufacturer],
        P.Placing_Basis, b.Controlling_Body AS Controlling_Body,
        CAST(CASE WHEN Coverage <> 'In-Orbit' THEN [Launch Date] ELSE P.Inception END AS DATE) AS On_Risk_Date,
        CAST(CASE WHEN Coverage <> 'In-Orbit' THEN DATEADD(MONTH, [MonthsOnRisk], IPO.[Launch Date]) ELSE P.Expiry END AS DATE) AS Off_Risk_Date,
        NULL AS [Action], Null as Altitude_km, Null as Line_size_perc
    FROM [Prequel_Reporting].[Prequel].[Layer_Details_v] P
    left JOIN SpacecraftData IPO ON P.LayerId = IPO.LayerId
    JOIN [SpaceTrax_Data].[rds].[param_mapping_code] PMC ON P.Mapping_Code = PMC.Mapping_Code
    join Prequel.Data.Layers_t l on l.LayerID=P.LayerID
    inner join Prequel.Lookups.Controlling_Bodies_t b
        on b.Controlling_BodyID=l.Controlling_BodyID and b.Controlling_Body in ('IG','MGU')
    inner join [SpaceTrax_Data].[rds].param_consortium_splits cs
        on cs.[Start Date]<=l.inception and cs.[End Date]>=l.inception
        AND cs.[Controlling Body] = 'MGU'
    WHERE P.IsBound = 1
    AND (CASE WHEN Coverage <> 'In-Orbit' THEN DATEADD(MONTH, [MonthsOnRisk], IPO.[Launch Date]) ELSE P.Expiry END) >= (SELECT TOP 1 [As_At_Date] FROM DateParam)
    AND P.Inception <= (SELECT TOP 1 [As_At_Date] FROM DateParam)
    and p.Mapping_Code in ('ASO','ASC')
    AND p.Program_Type NOT IN ('LVFO','Satellite Launch')
    AND p.Placing_Basis NOT IN ('Consortium Declaration')
),
ManualOverrides AS (
    SELECT
        ProgramId, LayerId, Program_Type, Underwriting_Status, Program_Name,
        Entity, Mapping_Code, [Bus Family], NULL As [Bus Type], [Vehicle Family],
        case when Consortium=1 and l.IGOnly<>1 then [Layer_Signed_Exposure_USD]*cs.[Cession] else [Layer_Signed_Exposure_USD] end as [Layer_Signed_Exposure_USD],
        [Layer_Signed_Premium_USD],
        Inception, Coverage, Expiry, [Launch Date], [Orbit Category],
        MonthsOnRisk, SpacecraftId, [Spacecraft Name], [Bus Manufacturer], [Prime Manufacturer],
        CAST(NULL AS NVARCHAR(100)) AS Placing_Basis,
        CAST(CASE WHEN l.IGOnly = 1 THEN 'IG' WHEN l.Consortium = 1 THEN 'Consortium' ELSE NULL END AS NVARCHAR(50)) AS Controlling_Body,
        On_Risk_Date, Off_Risk_Date, [Action], Altitude_km, Line_size_perc
    FROM [SpaceTrax_Data].[rds].[manually_controlled_rds_layers] l
    inner join [SpaceTrax_Data].[rds].param_consortium_splits cs
        on cs.[Start Date]<=isnull(l.Inception,getdate()) and cs.[End Date]>=isnull(l.Inception,getdate())
        and cs.[Controlling Body]='MGU'
    where ((l.Consortium=1 and l.LLoydsOnly<>1) or l.IGOnly=1)
),
BaseData AS (
    SELECT
        B.*,
        CASE WHEN DP.As_At_Date > B.On_Risk_Date AND B.Off_Risk_Date >= DP.As_At_Date THEN 1 ELSE 0 END AS On_Risk_Flag,
        RIC.QS_Cede AS QS_Cede,
        RIC.Notes AS RI_Notes,
        IG_QS.QS AS QS_IGR,
        XOL.Deductible_USD AS XOL_Deductible_USD,
        XOL.Limit_USD AS XOL_Limit_USD,
        CASE WHEN B.[Orbit Category] = 'LEO' THEN ROUND(((DATEDIFF(DAY, DP.As_At_Date, B.Off_Risk_Date) / 30.5) + 0.5), 0) ELSE 0 END AS MonthsLeftOnRisk
    FROM (
        SELECT * FROM ManualOverrides WHERE [Action] = 'Add Layer'
        union all
        select * from OriginalBaseDataConsortiumLayersAdjusted
        WHERE LayerID NOT IN (
            SELECT LayerID FROM [SpaceTrax_Data].[rds].[manually_controlled_rds_layers]
            WHERE ([Action] = 'Remove Layer' OR [Action] = 'Add Layer')
        )
        union all
        SELECT * FROM OriginalBaseDataIGLayersAdjusted
        where LayerID not in (select LayerID from OriginalBaseDataConsortiumLayersAdjusted)
        and LayerID NOT IN (SELECT LayerID FROM [SpaceTrax_Data].[rds].[manually_controlled_rds_layers])
    ) B
    CROSS JOIN DateParam DP
    LEFT JOIN [SpaceTrax_Data].[rds].[param_ri_cessions] RIC
        ON B.Inception BETWEEN RIC.Inception AND RIC.Expiry
    LEFT JOIN [SpaceTrax_Data].[rds].[param_ig_cessions_qs] IG_QS
        ON B.Entity = IG_QS.Entity AND B.Mapping_Code = IG_QS.Class_Code AND YEAR(B.Inception) = IG_QS.Year
    LEFT JOIN [SpaceTrax_Data].[rds].[param_ig_cessions_xol] XOL
        ON B.Entity = XOL.Entity AND DP.As_At_Date BETWEEN XOL.Inception AND XOL.Expiry
    WHERE B.Off_Risk_Date >= DP.As_At_Date
    AND DP.As_At_Date >= B.On_Risk_Date
),
QSCalculations AS (
    SELECT *,
        [Layer_Signed_Exposure_USD] * QS_Cede AS [External QS],
        [Layer_Signed_Exposure_USD] * (1- QS_Cede) AS [Net of External QS],
        [Layer_Signed_Exposure_USD] * (1- QS_Cede) * QS_IGR AS [QS FIBL Ceded],
        [Layer_Signed_Exposure_USD] * (1- QS_Cede) * (1 - QS_IGR) AS [Net of QS IGR]
    FROM BaseData B
),
XoLCalculations AS (
    SELECT FC.*,
        CASE WHEN FC.[Net of QS IGR] > FC.XOL_Deductible_USD THEN
            CASE WHEN (FC.[Net of QS IGR] - FC.XOL_Deductible_USD) < FC.XOL_Limit_USD
                 THEN (FC.[Net of QS IGR] - FC.XOL_Deductible_USD)
                 ELSE FC.XOL_Limit_USD END
        ELSE 0 END AS [XoL FIBL Ceded]
    FROM QSCalculations FC
)
SELECT
    ProgramId, LayerId, Program_Name, Entity, Mapping_Code, Inception, Expiry,
    Coverage, [Layer_Signed_Exposure_USD], [Layer_Signed_Premium_USD], [Launch Date],
    [Bus Type], [Bus Family], [Vehicle Family], [Orbit Category], MonthsOnRisk,
    SpacecraftId, [Spacecraft Name], [Bus Manufacturer], [Prime Manufacturer],
    Placing_Basis, Controlling_Body,
    On_Risk_Date, Off_Risk_Date, [Action], QS_IGR, XOL_Deductible_USD, XOL_Limit_USD,
    MonthsLeftOnRisk, [External QS], [Net of External QS], [QS FIBL Ceded],
    [Net of QS IGR], [XoL FIBL Ceded],
    [XoL FIBL Ceded] + [QS FIBL Ceded] AS [Total FIBL Ceded],
    [Net of QS IGR] - [XoL FIBL Ceded] AS [Net of XoL IGR]
    -- Renewal pointers are NOT joined here. They come from the OPTIONAL
    -- renewal-watch query (sql/renewal_watch.sql, run separately by
    -- renewals.load_watch), which degrades gracefully to the spacecraft
    -- heuristic if the PBI snapshot table is unavailable. Keeping that JOIN out
    -- of the core extract means a missing / renamed PBI table can never break
    -- the on-risk run — the renewal classification just falls back.
FROM XoLCalculations
