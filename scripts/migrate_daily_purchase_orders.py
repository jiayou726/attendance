"""Idempotently migrate a local SQLite database to one purchase order per day."""

from __future__ import annotations

import argparse
from collections import defaultdict
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
    connection.row_factory = sqlite3.Row
    try:
        connection.execute("BEGIN IMMEDIATE")
        columns = {row[1] for row in connection.execute("PRAGMA table_info(kitchen_purchase_order_item)")}
        if "supplier_id" not in columns:
            connection.execute("ALTER TABLE kitchen_purchase_order_item ADD COLUMN supplier_id INTEGER")
        if "supplier_name_snapshot" not in columns:
            connection.execute("ALTER TABLE kitchen_purchase_order_item ADD COLUMN supplier_name_snapshot VARCHAR(120)")
        connection.execute("""
            UPDATE kitchen_purchase_order_item
            SET supplier_id = (SELECT supplier_id FROM kitchen_purchase_order WHERE id = kitchen_purchase_order_item.order_id),
                supplier_name_snapshot = (SELECT supplier_name_snapshot FROM kitchen_purchase_order WHERE id = kitchen_purchase_order_item.order_id)
            WHERE supplier_name_snapshot IS NULL
        """)

        grouped = defaultdict(list)
        for order in connection.execute("SELECT id, service_date, status FROM kitchen_purchase_order ORDER BY service_date, id"):
            grouped[order["service_date"]].append(order)
        for service_date, orders in grouped.items():
            active = [row for row in orders if row["status"] != "cancelled"]
            included = active or orders
            primary = included[0]
            primary_id = primary["id"]
            included_ids = {row["id"] for row in included}
            status = "cancelled" if not active else ("confirmed" if all(row["status"] == "confirmed" for row in active) else "draft")
            seen = {
                row[0]
                for row in connection.execute(
                    "SELECT ingredient_id FROM kitchen_purchase_order_item WHERE order_id = ?", (primary_id,)
                )
            }
            for order in orders:
                if order["id"] == primary_id:
                    continue
                items = list(connection.execute(
                    "SELECT id, ingredient_id FROM kitchen_purchase_order_item WHERE order_id = ?", (order["id"],)
                ))
                for item in items:
                    if order["id"] not in included_ids or item["ingredient_id"] in seen:
                        connection.execute("DELETE FROM kitchen_purchase_order_item WHERE id = ?", (item["id"],))
                    else:
                        connection.execute(
                            "UPDATE kitchen_purchase_order_item SET order_id = ? WHERE id = ?", (primary_id, item["id"])
                        )
                        seen.add(item["ingredient_id"])
                connection.execute("DELETE FROM kitchen_purchase_order WHERE id = ?", (order["id"],))
            connection.execute("""
                UPDATE kitchen_purchase_order
                SET supplier_id = NULL, supplier_key = 'daily', supplier_name_snapshot = '每日採購單',
                    supplier_overridden = 0, status = ?
                WHERE id = ?
            """, (status, primary_id))
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
    print("每日採購單 migration 完成。")
    if backup:
        print(f"備份：{backup}")


if __name__ == "__main__":
    main()
