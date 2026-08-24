"""Add editable package conversion and ordered tracking to a local SQLite DB."""

from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
import shutil
import sqlite3


def migrate(db_path: Path, *, backup: bool = True) -> Path | None:
    db_path = db_path.resolve()
    if not db_path.is_file():
        raise SystemExit(f"找不到資料庫：{db_path}")
    backup_path = None
    if backup:
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        backup_path = db_path.with_name(f"{db_path.name}.backup-{stamp}")
        shutil.copy2(db_path, backup_path)

    connection = sqlite3.connect(db_path)
    try:
        connection.execute("BEGIN IMMEDIATE")
        columns = {row[1] for row in connection.execute("PRAGMA table_info(kitchen_purchase_order_item)")}
        if "package_qty" not in columns:
            connection.execute("ALTER TABLE kitchen_purchase_order_item ADD COLUMN package_qty NUMERIC(16, 4)")
        if "package_unit" not in columns:
            connection.execute("ALTER TABLE kitchen_purchase_order_item ADD COLUMN package_unit VARCHAR(20)")
        if "ordered" not in columns:
            connection.execute("ALTER TABLE kitchen_purchase_order_item ADD COLUMN ordered BOOLEAN NOT NULL DEFAULT 0")
        connection.commit()
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()
    return backup_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("database", type=Path)
    parser.add_argument("--no-backup", action="store_true")
    args = parser.parse_args()
    backup = migrate(args.database, backup=not args.no_backup)
    print("採購包裝換算與叫貨追蹤 migration 完成。")
    if backup:
        print(f"備份：{backup}")


if __name__ == "__main__":
    main()
