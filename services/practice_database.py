"""Isolated database bootstrap for the kitchen practice account."""
from __future__ import annotations

from datetime import datetime

from sqlalchemy import Column, DateTime, MetaData, String, Table, func, inspect, select, text

import ai_models
from extensions import db


# Copy only reusable master/catalog data. Operational records such as menu
# plans, purchase orders, AI drafts and daily notes intentionally start empty,
# so the practice account behaves like a fresh workspace rather than a mirror
# of production activity.
PRACTICE_SEED_TABLES = (
    "kitchen_school",
    "kitchen_supplier",
    "kitchen_ingredient",
    "kitchen_supplier_item",
    "kitchen_recipe",
    "kitchen_recipe_ingredient",
    "kitchen_recipe_tag",
    "kitchen_meal_profile",
)

_META = MetaData()
_PRACTICE_META = Table(
    "kitchen_practice_meta",
    _META,
    Column("key", String(80), primary_key=True),
    Column("value", String(255), nullable=False, default=""),
    Column("updated_at", DateTime, nullable=False, default=datetime.utcnow),
)
_SEED_KEY = "master_seed_v1"


def _database_identity(engine) -> str:
    return engine.url.render_as_string(hide_password=False)


def _ensure_additive_compatibility(engine) -> None:
    """Bring an older practice DB up to the current additive kitchen schema."""
    inspector = inspect(engine)
    with engine.begin() as connection:
        if inspector.has_table("kitchen_menu_assignment"):
            columns = {column["name"] for column in inspector.get_columns("kitchen_menu_assignment")}
            if "service_status" not in columns:
                connection.execute(text(
                    "ALTER TABLE kitchen_menu_assignment "
                    "ADD COLUMN service_status VARCHAR(20) NOT NULL DEFAULT 'serving'"
                ))

        if inspector.has_table("kitchen_school"):
            columns = {column["name"] for column in inspector.get_columns("kitchen_school")}
            if "default_vegetarian_headcount" not in columns:
                connection.execute(text(
                    "ALTER TABLE kitchen_school ADD COLUMN "
                    "default_vegetarian_headcount INTEGER NOT NULL DEFAULT 0"
                ))

        if inspector.has_table("kitchen_daily_dish_note"):
            columns = {column["name"] for column in inspector.get_columns("kitchen_daily_dish_note")}
            for column_name in ("combo_count", "class_count", "bento_count", "small_bento_count"):
                if column_name not in columns:
                    connection.execute(text(
                        f"ALTER TABLE kitchen_daily_dish_note ADD COLUMN {column_name} INTEGER"
                    ))
            if "school_class_counts" not in columns:
                connection.execute(text(
                    "ALTER TABLE kitchen_daily_dish_note "
                    "ADD COLUMN school_class_counts TEXT NOT NULL DEFAULT '{}'"
                ))

    ai_models.ensure_schema(engine)


def _advance_postgres_sequences(engine) -> None:
    if engine.dialect.name != "postgresql":
        return
    with engine.begin() as connection:
        for table_name in PRACTICE_SEED_TABLES:
            table = db.metadata.tables.get(table_name)
            if table is None or "id" not in table.c:
                continue
            max_id = connection.execute(select(func.max(table.c.id))).scalar()
            if max_id is None:
                continue
            # Table names come only from the constant tuple above.
            connection.execute(text(
                "SELECT setval(pg_get_serial_sequence(:table_name, 'id'), :max_id, true)"
            ), {"table_name": table_name, "max_id": int(max_id)})


def _seed_master_data(main_engine, practice_engine) -> None:
    _PRACTICE_META.create(bind=practice_engine, checkfirst=True)
    with practice_engine.connect() as connection:
        seeded = connection.execute(
            select(_PRACTICE_META.c.key).where(_PRACTICE_META.c.key == _SEED_KEY)
        ).first()
    if seeded:
        return

    main_tables = set(inspect(main_engine).get_table_names())
    with main_engine.connect() as source, practice_engine.begin() as target:
        for table_name in PRACTICE_SEED_TABLES:
            table = db.metadata.tables.get(table_name)
            if table is None or table_name not in main_tables:
                continue
            # A partially prepared practice DB is left intact instead of
            # overwriting user-created sandbox data.
            existing = target.execute(select(func.count()).select_from(table)).scalar_one()
            if existing:
                continue
            rows = source.execute(select(table)).mappings().all()
            if rows:
                target.execute(table.insert(), [dict(row) for row in rows])

        target.execute(_PRACTICE_META.insert().values(
            key=_SEED_KEY,
            value="Copied reusable kitchen master data from formal DB",
            updated_at=datetime.utcnow(),
        ))

    _advance_postgres_sequences(practice_engine)


def ensure_practice_database() -> None:
    """Create/upgrade the sandbox and seed reusable master data once.

    The function is called before the practice session flag is enabled, so the
    source engine is always the formal database. It fails closed if both binds
    accidentally point at the same database.
    """
    practice_engine = db.engines.get("practice")
    if practice_engine is None:
        raise RuntimeError("Practice database bind is not configured.")

    main_engine = db.engine
    if _database_identity(main_engine) == _database_identity(practice_engine):
        raise RuntimeError("Practice database must not be the formal database.")

    db.metadata.create_all(bind=practice_engine)
    _ensure_additive_compatibility(practice_engine)
    _seed_master_data(main_engine, practice_engine)
