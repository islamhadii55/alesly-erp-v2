#!/usr/bin/env python3
"""Create a consistent SQLite backup using SQLite's online backup API."""
import os
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

source = Path(os.environ.get("DATABASE_PATH", "original_auto.db"))
backup_dir = Path(os.environ.get("BACKUP_DIR", source.parent / "backups"))
backup_dir.mkdir(parents=True, exist_ok=True)
if not source.exists():
    raise SystemExit(f"Database not found: {source}")

timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
destination = backup_dir / f"original_auto-{timestamp}.db"
with sqlite3.connect(f"file:{source}?mode=ro", uri=True) as source_conn:
    with sqlite3.connect(destination) as destination_conn:
        source_conn.backup(destination_conn)
        destination_conn.execute("PRAGMA integrity_check")
print(destination)
