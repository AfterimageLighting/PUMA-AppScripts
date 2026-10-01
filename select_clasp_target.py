#!/usr/bin/env python3
"""Select the explicit Apps Script deployment target for the current checkout."""

import json
import shutil
import sys
from pathlib import Path

TARGETS = {
    "test": Path(".clasp.test.json"),
    "production": Path(".clasp.production.json"),
    "prod": Path(".clasp.production.json"),
}

if len(sys.argv) != 2 or sys.argv[1].lower() not in TARGETS:
    raise SystemExit("Usage: python select_clasp_target.py test|production")

name = sys.argv[1].lower()
source = TARGETS[name]
if not source.exists():
    raise SystemExit(f"Missing target config: {source}")

obj = json.loads(source.read_text(encoding="utf-8"))
script_id = str(obj.get("scriptId", "")).strip()
if not script_id:
    raise SystemExit(f"{source} does not contain a Script ID")

shutil.copyfile(source, ".clasp.json")
print(f"Selected {name} Apps Script target: ...{script_id[-10:]}")
print("Run 'clasp push' only after confirming the intended environment.")
