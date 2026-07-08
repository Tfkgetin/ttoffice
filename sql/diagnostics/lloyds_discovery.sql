/* ===================================================================
   Lloyd's (S3123) RDS — discovery queries
   Run each block in SSMS against SpaceTrax_Data (LON-SQLP-V005) and
   paste the results back. These unlock: (a) JJ's exact RDS logic,
   (b) the 'BJ-3C 01' CAST bug, (c) the bus-type mapping, (d) altitude.
   =================================================================== */

/* -------------------------------------------------------------------
   1) THE PRIZE — the two Lloyd's view definitions (JJ's actual logic).
      Paste both results verbatim. From these I can replicate every RDS
      in the pipeline to the dollar (and the data team can see the CAST).
   ------------------------------------------------------------------- */
SELECT OBJECT_DEFINITION(OBJECT_ID('rds.vw_SpaceRDS_All_Lloyds_RDS'))      AS all_lloyds_rds_sql;
SELECT OBJECT_DEFINITION(OBJECT_ID('rds.vw_SpaceRDS_SpaceWeather_Lloyds')) AS spaceweather_lloyds_sql;

-- any other Lloyd's views the two above build on (they're often a UNION
-- of per-scenario views):
SELECT s.name + '.' + v.name AS view_name
FROM sys.views v JOIN sys.schemas s ON s.schema_id = v.schema_id
WHERE v.name LIKE '%Lloyd%'
ORDER BY view_name;

/* -------------------------------------------------------------------
   2) THE CAST BUG — the varchar 'BJ-3C 01' that dies converting to int.
      This finds every spacecraft whose Seradata id is non-numeric (so a
      CAST(... AS INT) in the view blows up). 'BJ-3C 01' should appear.
      Fix = the view should TRY_CAST, or these ids need cleaning.
   ------------------------------------------------------------------- */
SELECT [Seradata Spacecraft ID], [Spacecraft Name], [Bus Type], [Bus Manufacturer]
FROM SpaceTrax_Data.dbo.tbl_SpaceCraft
WHERE [Seradata Spacecraft ID] IS NOT NULL
  AND TRY_CONVERT(INT, [Seradata Spacecraft ID]) IS NULL;

/* -------------------------------------------------------------------
   3) BUS-TYPE MAPPING — Seradata [Bus Type] -> Lloyd's bus-type group.
      The main extract already carries [Bus Type]; I just need this map to
      roll it up to JJ's groups for Space Weather – Design Deficiency.
   ------------------------------------------------------------------- */
SELECT * FROM SpaceTrax_Data.rds.params_Satellite_List_Lloyds;

/* -------------------------------------------------------------------
   4) ALTITUDE — for Space Debris "worst LEO orbit range".
   4a) the LEO altitude-group bands (Group 1 = 600±200, Group 2 = 1400±200):
   ------------------------------------------------------------------- */
SELECT * FROM SpaceTrax_Data.rds.params_lloyds_proton_flare_altitude;

/* 4b) where altitude lives on the spacecraft table (apogee/perigee/etc.) */
SELECT COLUMN_NAME, DATA_TYPE
FROM SpaceTrax_Data.INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'tbl_SpaceCraft'
  AND (COLUMN_NAME LIKE '%apogee%'  OR COLUMN_NAME LIKE '%perigee%'
    OR COLUMN_NAME LIKE '%altitude%' OR COLUMN_NAME LIKE '%height%'
    OR COLUMN_NAME LIKE '%orbit%'    OR COLUMN_NAME LIKE '%semi%major%');

/* -------------------------------------------------------------------
   5) RISK APPETITE — the threshold behind Breaches_Risk_Appetite /
      "> Risk Appetite USD?" (so I can wire the breach flags).
      Guessing the table name; adjust if it errors.
   ------------------------------------------------------------------- */
SELECT * FROM SpaceTrax_Data.rds.param_lloyds_risk_appetite;   -- if this name is wrong:
-- SELECT s.name+'.'+t.name FROM sys.tables t JOIN sys.schemas s ON s.schema_id=t.schema_id
-- WHERE t.name LIKE '%appetite%' OR t.name LIKE '%risk_app%';
