"""Export all SQLite tables to JSON for migration verification."""
import json
import sqlite3
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DB = ROOT / "database" / "bam_motor_group.db"
OUT = ROOT / "backups" / f"sqlite_export_{datetime.now():%Y%m%d_%H%M%S}.json"

conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row
tables = [r[0] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' ORDER BY name")]
payload = {"database": str(DB), "exported_at": datetime.now().isoformat(), "tables": {}}
for table in tables:
    rows = conn.execute(f'SELECT * FROM "{table}"').fetchall()
    payload["tables"][table] = [dict(r) for r in rows]
conn.close()
OUT.parent.mkdir(parents=True, exist_ok=True)
OUT.write_text(json.dumps(payload, indent=2, default=str), encoding="utf-8")
print(OUT)
