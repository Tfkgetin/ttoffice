#!/usr/bin/env python3
"""SQL connection tester for the Space RDS pipeline.

Connects to the view, prints columns, row count, sample rows, and a
column-mapping preview — WITHOUT running the pipeline. Use this first.

Usage:
    python test_sql.py --server LON-SQL-XXX --database SpaceTrax_Data --view rds.vw_SpaceRDS_OnRisk
"""
from __future__ import annotations
import argparse
import re
import sys

try:
    import pyodbc
except ImportError:
    sys.exit("pyodbc not installed — run:  pip install pyodbc")


# Pipeline-required fields, keyed by normalised SQL column name
TARGETS = {
    "programid": "program_id", "layerid": "layer_id", "coverage": "coverage",
    "monthsonrisk": "months_on_risk", "spacecraftid": "spacecraft_id",
    "launchdate": "launch_date", "inception": "inception", "expiry": "expiry",
    "orbitcategory": "orbit", "orbit": "orbit",
    "onriskdate": "on_risk_date", "offriskdate": "off_risk_date",
    "busmanufacturer": "bus_manufacturer", "primemanufacturer": "prime_manufacturer",
    "programtype": "program_type", "underwritingstatus": "underwriting_status",
    "programname": "program_name", "entity": "entity", "mappingcode": "mapping_code",
    "spacecraftname": "spacecraft_name", "busfamily": "bus_family",
    "vehiclefamily": "vehicle_family", "layersignedexposure": "layer_signed_exposure",
    "isconsortium": "is_consortium",
}
REQUIRED = set(TARGETS.values()) - {"is_consortium"}  # is_consortium nice-to-have


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


def pick_driver() -> str:
    drivers = pyodbc.drivers()
    print(f"Installed ODBC drivers: {drivers}")
    for want in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server",
                 "SQL Server Native Client 11.0", "SQL Server"):
        if want in drivers:
            return want
    sys.exit("No SQL Server ODBC driver found — install 'ODBC Driver 18 for SQL Server'.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--server", required=True)
    ap.add_argument("--database", required=True)
    ap.add_argument("--view", required=True)
    args = ap.parse_args()

    driver = pick_driver()
    conn_str = (f"DRIVER={{{driver}}};SERVER={args.server};"
                f"DATABASE={args.database};Trusted_Connection=yes;")
    if "18" in driver:
        conn_str += "TrustServerCertificate=yes;"

    print(f"\nConnecting: {conn_str}")
    cn = pyodbc.connect(conn_str, timeout=15)
    cur = cn.cursor()

    cur.execute(f"SELECT COUNT(*) FROM {args.view}")
    n = cur.fetchone()[0]
    print(f"Connected. {args.view} has {n} rows.\n")

    cur.execute(f"SELECT TOP 3 * FROM {args.view}")
    cols = [d[0] for d in cur.description]
    print(f"Columns ({len(cols)}):")
    mapped, extra = {}, []
    for c in cols:
        t = TARGETS.get(norm(c))
        if t:
            mapped[t] = c
            print(f"  {c:35s} -> {t}")
        else:
            extra.append(c)
            print(f"  {c:35s} -> (unmapped — ignored)")

    missing = REQUIRED - set(mapped)
    print()
    if missing:
        print(f"MISSING required fields: {sorted(missing)}")
        print("Send this whole output back and the mapping will be extended.")
    else:
        print("All required fields mapped ✓")
    if "is_consortium" not in mapped:
        print("NOTE: no Is Consortium column — pipeline will need it from "
              "another source (it drives equity share + S3123 eligibility).")

    print("\nSample rows:")
    for row in cur.fetchall():
        print(" ", str(row)[:200])

    cn.close()


if __name__ == "__main__":
    main()
