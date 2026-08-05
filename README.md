# BAM Dealer Enterprise v16.0

One system for assets, parts, workshop, shipping, sales, tooling and business records.

## Start
Double-click `START_BAM.bat`.

## Upgrade existing data
Use `IMPORT_OLD_DATABASE.bat`, or copy the old `database` and `uploads` folders while BAM is stopped. v16 automatically upgrades the database.

## New in v16
See `docs/V16_RELEASE_NOTES.md`.

# BAM Motor Group Dealer Management System

**BAM Motor Group — Buy • Sell • Trade**

This is the permanent project structure for the BAM Dealer application.

## Project folders

- `app/` — Flask application, templates and styling
- `database/` — the permanent `bam_motor_group.db`
- `uploads/` — vehicle photos, receipts and documents
- `reports/` — generated exports and reports
- `backups/` — database backups
- `docs/` — operating and development documentation
- `tests/` — project health checks
- `tools/` — database migration and maintenance utilities

## First installation

1. Extract the project to `C:\BAM\BAMDealer`.
2. Copy your current working `bam_motor_group.db` into the `database` folder,
   or double-click `IMPORT_OLD_DATABASE.bat`.
3. Copy your current `uploads` files into this project's `uploads` folder.
4. Double-click `START_BAM.bat`.
5. Open `http://127.0.0.1:5000`.

## Normal daily use

Double-click `START_BAM.bat`, keep the black window open and use the app in Edge.

## Backup

Use the Backup button inside the web app or double-click `BACKUP_BAM.bat`.

## Important

From now on, keep this one permanent project folder. Future upgrades should
replace application files only and must preserve the `database`, `uploads` and
`backups` folders.


## v13.3 mixed stock
Use Asset Type on Add Stock or Edit Stock to manage cars, caravans and boats. Existing records default to Car.


## v15.2 Assets & Equipment
Adds equipment inventory, tool photos and receipts, maintenance dates, check-in/check-out, assignments, locations and movement history.


## v15.3 Editable Invoices
Open **Invoice Centre** from the sidebar to search, view and edit saved vehicle sale invoices. Every update records a reason and history entry.


## v15.4 Deal Files
- Complete deal file for every recorded sale.
- Buyer, invoice, payment, trade-in and warranty summary.
- Delivery checklist and deal notes.
- Links to documents, invoice history, expenses and workshop records.


## v15.5 Sales & Pricing
- Minimum sale price and asking price are editable on the stock-item editor.
- Live margin estimates compare both prices with the purchase price.
- The prepare-for-sale pricing task completes automatically once both prices are saved.
