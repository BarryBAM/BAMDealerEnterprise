from pathlib import Path
import shutil
import sys
from datetime import datetime

PROJECT = Path(__file__).resolve().parents[1]
DEST = PROJECT / "database" / "bam_motor_group.db"
BACKUPS = PROJECT / "backups"
BACKUPS.mkdir(exist_ok=True)

if len(sys.argv) < 2:
    print('Usage: python tools\\import_existing_database.py "C:\\path\\to\\bam_motor_group.db"')
    raise SystemExit(1)

source = Path(sys.argv[1]).expanduser().resolve()
if not source.exists():
    print(f"Database not found: {source}")
    raise SystemExit(1)

if DEST.exists():
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = BACKUPS / f"bam_motor_group_before_import_{timestamp}.db"
    shutil.copy2(DEST, backup)
    print(f"Current project database backed up to: {backup}")

shutil.copy2(source, DEST)
print(f"Database imported successfully to: {DEST}")
