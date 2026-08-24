"""Add and backfill supplier catalog links on procurement items."""

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
        if "supplier_item_id" not in columns:
            connection.execute("ALTER TABLE kitchen_purchase_order_item ADD COLUMN supplier_item_id INTEGER")
        if "package_conversion_snapshot" not in columns:
            connection.execute(
                "ALTER TABLE kitchen_purchase_order_item ADD COLUMN package_conversion_snapshot VARCHAR(120)"
            )
        connection.execute("""
            UPDATE kitchen_purchase_order_item AS poi
            SET supplier_item_id = (
                SELECT si.id
                FROM kitchen_supplier_item AS si
                WHERE si.supplier_id = poi.supplier_id
                  AND si.active = 1
                  AND (si.ingredient_id = poi.ingredient_id OR lower(trim(si.name)) = lower(trim(poi.ingredient_name_snapshot)))
                ORDER BY CASE WHEN si.ingredient_id = poi.ingredient_id THEN 0 ELSE 1 END, si.id
                LIMIT 1
            )
            WHERE poi.supplier_id IS NOT NULL
        """)
        connection.execute("""
            UPDATE kitchen_purchase_order_item
            SET package_conversion_snapshot = (
                SELECT package_conversion
                FROM kitchen_supplier_item
                WHERE id = kitchen_purchase_order_item.supplier_item_id
            )
            WHERE supplier_item_id IS NOT NULL
        """)
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
    print("採購單已連結廠商品項換算資料。")
    if backup:
        print(f"備份：{backup}")


if __name__ == "__main__":
    main()
