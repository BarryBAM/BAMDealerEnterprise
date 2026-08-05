# Upgrade Guide

Future BAM upgrades should follow this process:

1. Stop BAM with `Ctrl + C`.
2. Run `BACKUP_BAM.bat`.
3. Preserve these folders:
   - `database`
   - `uploads`
   - `backups`
4. Replace only:
   - `app`
   - `requirements.txt`
   - launcher or tool files when supplied
5. Start BAM and run `RUN_HEALTH_CHECK.bat`.

Never delete the permanent database unless a verified backup exists.
