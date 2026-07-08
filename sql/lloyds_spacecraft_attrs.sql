-- Per-spacecraft Lloyd's attributes for the S3123 RDS, keyed by the Seradata
-- Spacecraft ID so it joins the pipeline's per-layer frame directly:
--   * LloydsBusType — the Lloyd's bus-type group, from the Seradata
--     (Bus Type, Bus Manufacturer) -> group map (params_Satellite_List_Lloyds).
--     Drives Space Weather - Design Deficiency (worst bus type, top-4).
--   * AltitudeKm — latest altitude, for the Space Debris LEO orbit-range
--     grouping (Group 1 = 400-800 km, Group 2 = 1200-1600 km).
-- LEFT JOIN so altitude is returned even when a spacecraft's bus type is not in
-- the Lloyd's list. Avoids the vw_SpaceRDS_*_Lloyds views entirely (those error
-- on a varchar->int CAST), so it is robust to that bug.
SELECT sc.[Seradata Spacecraft ID]         AS SpacecraftId,
       sl.[Lloyds Satellite Bus Type List] AS LloydsBusType,
       sc.[Altitude Latest (km)]           AS AltitudeKm
FROM   SpaceTrax_Data.dbo.tbl_SpaceCraft sc
LEFT JOIN SpaceTrax_Data.rds.params_Satellite_List_Lloyds sl
       ON sl.[Bus Type]         = sc.[Bus Type]
      AND sl.[Bus Manufacturer] = sc.[Bus Manufacturer];
