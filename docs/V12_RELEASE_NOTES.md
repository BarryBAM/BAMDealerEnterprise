# BAM Dealer Enterprise v12.0

## Included in this baseline

- Repaired the Python indentation/startup error in `init_db()`.
- Preserved the existing vehicle, sales, workshop, advertising and reporting features.
- Preserved the existing SQLite database, uploads, reports and backups.
- Extended the existing Parts Inventory table instead of replacing it.
- Added safe database columns for source vehicle, VIN, make, model, year, condition,
  selling price, status, engine code, transmission code, barcode and date added.
- Added `part_photos` and `part_sales` tables for the next Parts Centre stages.
- Added missing `csv` and `io` imports and the `BASE_DIR` project path used by exports/imports.
- Kept a copy of the broken pre-v12 file as `app/app_pre_v12_broken_backup.py`.

## Start BAM Dealer

Double-click `START_BAM.bat`. The launcher creates/uses the local Python environment,
installs requirements and starts the system at `http://127.0.0.1:5000`.

## Existing data

The included `database/bam_motor_group.db` is retained. Before replacing an older copy,
keep a backup of its `database` and `uploads` folders.
