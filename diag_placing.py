#!/usr/bin/env python3
"""Diagnostic: what does ingest actually return for placing_basis?"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from src.parameters import Params
from src import ingest

p = Params.load(sys.argv[sys.argv.index("--config")+1])
if "--as-at" in sys.argv:
    import datetime as dt
    p.as_at = dt.date.fromisoformat(sys.argv[sys.argv.index("--as-at")+1])

df = ingest.load(p)
print(f"Rows: {len(df)}")
print(f"Columns: {list(df.columns)}")
print(f"\n'placing_basis' in columns: {'placing_basis' in df.columns}")
if "placing_basis" in df.columns:
    print("\nplacing_basis value counts:")
    print(df["placing_basis"].astype(str).value_counts().to_string())
    # show the 6 aggregate layers specifically
    agg = df[df["program_id"].astype(str).isin(["379919","361595","348005"])]
    print(f"\nAggregate-program rows ({len(agg)}):")
    for _, r in agg.iterrows():
        print(f"  {r['program_id']}/{r['layer_id']}  placing_basis='{r.get('placing_basis')}'")
else:
    print("\n>>> placing_basis NOT mapped. Check: (1) SQL SELECT outputs Placing_Basis,")
    print(">>> (2) ingest.py _SQL_TARGETS has 'placingbasis': 'placing_basis'.")
    print(f"\nConfig exclude_placing_basis = {p.raw.get('exclude_placing_basis')}")
