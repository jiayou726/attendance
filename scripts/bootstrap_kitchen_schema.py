"""Create the final kitchen_* schema without touching attendance tables.

Use this only for an initial deployment where production has no kitchen tables yet.
If any existing kitchen table is missing expected columns, this script refuses to alter it.
That fail-closed behavior prevents accidental partial upgrades on production.
"""

from sqlalchemy import inspect

from app import create_app
from extensions import db


KITCHEN_TABLES = (
    "kitchen_school",
    "kitchen_supplier",
    "kitchen_ingredient",
    "kitchen_supplier_item",
    "kitchen_recipe",
    "kitchen_recipe_ingredient",
    "kitchen_menu_plan",
    "kitchen_menu_plan_item",
    "kitchen_menu_assignment",
    "kitchen_purchase_order",
    "kitchen_purchase_order_item",
)

FORBIDDEN_TABLES = {"employee", "checkin", "employees", "checkins"}


def main():
    app = create_app({"AUTO_CREATE_DB": False})
    with app.app_context():
        engine = db.engine
        inspector = inspect(engine)
        existing = set(inspector.get_table_names())

        # This script must never mutate attendance tables.
        assert not (set(KITCHEN_TABLES) & FORBIDDEN_TABLES)

        mismatches = []
        for name in KITCHEN_TABLES:
            if name not in existing:
                continue
            actual_columns = {column["name"] for column in inspector.get_columns(name)}
            expected_columns = set(db.metadata.tables[name].columns.keys())
            missing = sorted(expected_columns - actual_columns)
            if missing:
                mismatches.append((name, missing))

        if mismatches:
            print("REFUSING TO ALTER EXISTING KITCHEN TABLES.")
            for table, missing in mismatches:
                print(f"- {table}: missing columns {', '.join(missing)}")
            print("Use the reviewed kitchen migration on a staging clone first.")
            raise SystemExit(2)

        created = []
        for name in KITCHEN_TABLES:
            if name in existing:
                continue
            table = db.metadata.tables[name]
            table.create(bind=engine, checkfirst=True)
            created.append(name)

        if created:
            print("Created kitchen tables:")
            for name in created:
                print(f"- {name}")
        else:
            print("Kitchen schema already matches the expected columns; nothing created.")

        print("Attendance tables were not altered.")


if __name__ == "__main__":
    main()
