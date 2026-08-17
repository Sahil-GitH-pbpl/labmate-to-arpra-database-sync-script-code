# Labmate to Arpra database sync report

Windows automation that restores the newest Labmate SQL Server `.bak` and reports rows that exist in SQL Server but are missing from the configured Arpra MySQL database.

## Behaviour

- Selects the newest stable `.bak` from `BackupFolder`.
- Detects an already processed backup by filename. A duplicate skips restore but still runs comparison.
- Copies a new backup to the SQL working folder, runs `RESTORE VERIFYONLY`, then restores `Bhasin_7001` after confirmation.
- Compares 14 configured tables against `10.1.1.67 / bhasin_7001_new` using configured keys.
- Reports only SQL-only/new rows. Existing row values and MySQL-only rows are ignored.
- Filters `address` comparison to `Atype C` and `D` only.
- Adds a `CompCatID` count summary for new `panelrates` rows.
- Never inserts, updates, or deletes data in target MySQL.

## Run

Place the backup in the configured backup folder, open a terminal in the script directory, and run:

```cmd
py run_labmate_backup.py
```

Compare without restoring:

```cmd
py run_labmate_backup.py --compare-only
```

Reports and live launcher logs are written to `ReportFolder`.

## Credentials

No password is committed to this repository. Set the MySQL password in one of these ways:

1. Set the environment variable configured by `PasswordEnvironmentVariable`:

   ```cmd
   setx LABMATE_MYSQL_PASSWORD "your-password"
   ```

2. Create a Windows-user DPAPI encrypted password file named `remote-mysql-password.dpapi` next to the scripts. The file is ignored by Git.

SQL Server defaults to Windows Authentication. Adjust `labmate-config.json` for another environment.

## Safety

- Restore requires administrator rights and typed `RESTORE` confirmation unless `--force` is deliberately used.
- Validation happens before the current SQL database is overwritten.
- Runtime reports, backup files, processed state, and credentials are excluded from Git.
