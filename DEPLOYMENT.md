# Deployment and Operations

## Required Railway variables

Set `APP_SECRET_KEY` to a random value of at least 32 characters. Set `INITIAL_ADMIN_PASSWORD` only for the first boot of an empty database, then remove it from the service variables. Keep `DATABASE_PATH=/data/original_auto.db` and attach a persistent Railway Volume mounted at `/data`.

The application refuses to start in production when `APP_SECRET_KEY` is missing or too short. It also refuses to initialize an empty production database without `INITIAL_ADMIN_PASSWORD`, preventing deployment with a known default password.

## Health checks

Railway uses `/health`. The endpoint checks both the Flask process and a basic SQLite query, returning HTTP 503 when the database cannot be opened. It intentionally does not expose internal exception details.

## Backups

Run the saved backup utility from the service environment or a trusted operator machine:

```bash
DATABASE_PATH=/data/original_auto.db BACKUP_DIR=/data/backups python scripts/backup_db.py
```

Copy encrypted backup files to storage outside the Railway Volume and periodically test restoration into a separate SQLite file. The backup utility uses SQLite's online backup API rather than copying a live database file directly.

## Verification before deployment

```bash
python -m unittest discover -s tests -v
python -m py_compile app.py
```

The deployment command is intentionally identical in `Procfile`, `nixpacks.toml`, and `railway.toml`. Dependencies are pinned in `requirements.txt` so a future package release cannot silently change the build.
