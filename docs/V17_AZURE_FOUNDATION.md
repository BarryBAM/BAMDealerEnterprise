# BAM Dealer Enterprise Cloud v17.0 — Azure Foundation

This release is the first deployable Azure foundation. It keeps the existing SQLite database so the current BAM data model and features remain intact while cloud access is tested.

## What changed

- Production startup with Gunicorn.
- Azure persistent storage through `BAM_DATA_ROOT` (recommended: `/home/data`).
- Secrets and operational settings moved to environment variables.
- Secure-cookie, proxy and session configuration for HTTPS behind Azure App Service.
- `/health` and `/ready` endpoints.
- Dockerfile, `startup.sh`, `azure.yaml`, `.env.example` and `.gitignore`.
- Local Windows operation remains available through `START_BAM.bat`.
- SQLite JSON export tool for migration checking.

## Important architecture limit

This foundation must run as **one App Service instance** because SQLite is a single-file database. Do not enable scale-out. PostgreSQL conversion is the next migration stage before multiple web instances or heavier simultaneous use.

## Azure App Service settings

Set these Application Settings:

- `BAM_SECRET_KEY`: a long random value.
- `BAM_DATA_ROOT`: `/home/data`.
- `BAM_SECURE_COOKIES`: `1`.
- `BAM_SESSION_HOURS`: `12`.
- `BAM_MAX_UPLOAD_MB`: `100`.
- `SCM_DO_BUILD_DURING_DEPLOYMENT`: `true`.

Startup Command:

```bash
./startup.sh
```

Health check path:

```text
/health
```

## Moving existing data

Copy the current database to `/home/data/database/bam_motor_group.db` and uploaded files to `/home/data/uploads/`. Keep the original local project untouched until cloud verification is complete.
