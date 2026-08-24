"""Make one purchase order per service date and snapshot supplier per item.

Revision ID: kitchen20260826
Revises: kitchen20260825
"""

from collections import defaultdict

from alembic import op
import sqlalchemy as sa


revision = "kitchen20260826"
down_revision = "kitchen20260825"
branch_labels = None
depends_on = None


def _columns(table):
    inspector = sa.inspect(op.get_bind())
    if not inspector.has_table(table):
        return set()
    return {column["name"] for column in inspector.get_columns(table)}


def upgrade():
    columns = _columns("kitchen_purchase_order_item")
    if "supplier_id" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column("supplier_id", sa.Integer(), nullable=True))
    if "supplier_name_snapshot" not in columns:
        op.add_column("kitchen_purchase_order_item", sa.Column("supplier_name_snapshot", sa.String(120), nullable=True))

    bind = op.get_bind()
    bind.execute(sa.text("""
        UPDATE kitchen_purchase_order_item
        SET supplier_id = (SELECT supplier_id FROM kitchen_purchase_order WHERE id = kitchen_purchase_order_item.order_id),
            supplier_name_snapshot = (SELECT supplier_name_snapshot FROM kitchen_purchase_order WHERE id = kitchen_purchase_order_item.order_id)
        WHERE supplier_name_snapshot IS NULL
    """))

    orders = bind.execute(sa.text("""
        SELECT id, service_date, status
        FROM kitchen_purchase_order
        ORDER BY service_date, id
    """)).mappings().all()
    by_date = defaultdict(list)
    for order in orders:
        by_date[order["service_date"]].append(order)

    for service_date, day_orders in by_date.items():
        active = [row for row in day_orders if row["status"] != "cancelled"]
        included = active or day_orders
        primary = included[0]
        primary_id = primary["id"]
        included_ids = {row["id"] for row in included}
        status = "cancelled" if not active else ("confirmed" if all(row["status"] == "confirmed" for row in active) else "draft")

        seen = {
            row[0]
            for row in bind.execute(sa.text(
                "SELECT ingredient_id FROM kitchen_purchase_order_item WHERE order_id = :order_id"
            ), {"order_id": primary_id})
        }
        for row in day_orders:
            if row["id"] == primary_id:
                continue
            items = bind.execute(sa.text("""
                SELECT id, ingredient_id FROM kitchen_purchase_order_item WHERE order_id = :order_id
            """), {"order_id": row["id"]}).mappings().all()
            for item in items:
                if row["id"] not in included_ids or item["ingredient_id"] in seen:
                    bind.execute(sa.text("DELETE FROM kitchen_purchase_order_item WHERE id = :id"), {"id": item["id"]})
                else:
                    bind.execute(sa.text("UPDATE kitchen_purchase_order_item SET order_id = :primary WHERE id = :id"), {"primary": primary_id, "id": item["id"]})
                    seen.add(item["ingredient_id"])
            bind.execute(sa.text("DELETE FROM kitchen_purchase_order WHERE id = :id"), {"id": row["id"]})

        bind.execute(sa.text("""
            UPDATE kitchen_purchase_order
            SET supplier_id = NULL,
                supplier_key = 'daily',
                supplier_name_snapshot = '每日採購單',
                supplier_overridden = :overridden,
                status = :status
            WHERE id = :id
        """), {"overridden": False, "status": status, "id": primary_id})


def downgrade():
    # Non-destructive: a daily order cannot safely be split back into supplier orders.
    pass
