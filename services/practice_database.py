"""Named, isolated kitchen practice workspaces.

PostgreSQL workspaces have private writable copies of master and operational
tables. Existing view-based accounts upgrade in place without erasing records.
SQLite uses a file-per-account fallback for development and tests.
"""
from __future__ import annotations

import re
import unicodedata
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path

from flask import current_app, session
from sqlalchemy import MetaData, create_engine, func, inspect, select, text
from sqlalchemy.exc import IntegrityError
from sqlalchemy.schema import CreateIndex, CreateTable, ForeignKeyConstraint

import ai_models  # noqa: F401 - register additive kitchen models
from extensions import db


MASTER_TABLES = (
    "kitchen_school", "kitchen_supplier", "kitchen_ingredient",
    "kitchen_supplier_item", "kitchen_recipe", "kitchen_recipe_ingredient",
    "kitchen_recipe_tag", "kitchen_meal_profile",
)
OPERATIONAL_TABLES = (
    "kitchen_menu_plan", "kitchen_menu_plan_item", "kitchen_daily_dish_note",
    "kitchen_menu_assignment", "kitchen_purchase_order",
    "kitchen_purchase_order_item", "kitchen_menu_draft",
    "kitchen_menu_draft_item",
)
_CONTROL_SCHEMA = "practice_control"
_CONTROL_TABLE = "account"
_WORKSPACE_RE = re.compile(r"^practice_[0-9a-f]{24}$")


class PracticeAccountError(RuntimeError):
    pass


class PracticeCapacityError(PracticeAccountError):
    pass


class PracticeNameError(PracticeAccountError):
    pass


def _utcnow() -> datetime:
    return datetime.now(timezone.utc)


def _normalize_name(value: str) -> tuple[str, str]:
    display = " ".join(unicodedata.normalize("NFKC", value or "").split())
    if not 2 <= len(display) <= 30:
        raise PracticeNameError("姓名請輸入 2～30 個字。")
    return display, display.casefold()


def _practice_engine():
    engine = db.engines.get("practice")
    if engine is None:
        raise PracticeAccountError("練習資料庫尚未設定。")
    return engine


def _is_postgres(engine) -> bool:
    return engine.dialect.name == "postgresql"


def _control_name(engine) -> str:
    return f"{_CONTROL_SCHEMA}.{_CONTROL_TABLE}" if _is_postgres(engine) else "practice_account"


def _ensure_control(engine) -> None:
    with engine.begin() as connection:
        if _is_postgres(engine):
            connection.execute(text(f"CREATE SCHEMA IF NOT EXISTS {_CONTROL_SCHEMA}"))
            connection.execute(text(f"""
                CREATE TABLE IF NOT EXISTS {_CONTROL_SCHEMA}.{_CONTROL_TABLE} (
                    id VARCHAR(32) PRIMARY KEY,
                    display_name VARCHAR(60) NOT NULL,
                    normalized_name VARCHAR(120) NOT NULL UNIQUE,
                    workspace_key VARCHAR(63) NOT NULL UNIQUE,
                    created_at TIMESTAMPTZ NOT NULL,
                    last_active_at TIMESTAMPTZ NOT NULL
                )
            """))
            connection.execute(text(
                f"REVOKE ALL ON SCHEMA {_CONTROL_SCHEMA} FROM PUBLIC, anon, authenticated"
            ))
        else:
            connection.execute(text("""
                CREATE TABLE IF NOT EXISTS practice_account (
                    id VARCHAR(32) PRIMARY KEY,
                    display_name VARCHAR(60) NOT NULL,
                    normalized_name VARCHAR(120) NOT NULL UNIQUE,
                    workspace_key VARCHAR(63) NOT NULL UNIQUE,
                    created_at DATETIME NOT NULL,
                    last_active_at DATETIME NOT NULL
                )
            """))


def _row_dict(row):
    if row is None:
        return None
    data = dict(row._mapping)
    for key in ("created_at", "last_active_at"):
        value = data.get(key)
        if isinstance(value, str):
            data[key] = datetime.fromisoformat(value)
        if data.get(key) and data[key].tzinfo is None:
            data[key] = data[key].replace(tzinfo=timezone.utc)
    return data


def _account_by(engine, column: str, value: str, connection=None):
    if column not in {"id", "normalized_name"}:
        raise ValueError("Unsupported practice account lookup")
    query = text(
        f"SELECT id, display_name, normalized_name, workspace_key, created_at, last_active_at "
        f"FROM {_control_name(engine)} WHERE {column}=:value"
    )
    if connection is not None:
        return _row_dict(connection.execute(query, {"value": value}).first())
    with engine.connect() as local_connection:
        return _row_dict(local_connection.execute(query, {"value": value}).first())


def _quote(engine, identifier: str) -> str:
    return engine.dialect.identifier_preparer.quote(identifier)


def _assert_workspace_key(value: str) -> str:
    if not _WORKSPACE_RE.fullmatch(value or ""):
        raise PracticeAccountError("練習空間識別碼不正確。")
    return value


def _clone_operational_tables(connection, schema: str, names=OPERATIONAL_TABLES) -> None:
    """Create private tables without cross-schema foreign keys."""
    metadata = MetaData()
    for table_name in names:
        cloned = db.metadata.tables[table_name].to_metadata(metadata, schema=schema)
        for constraint in list(cloned.constraints):
            if isinstance(constraint, ForeignKeyConstraint):
                cloned.constraints.remove(constraint)
    # Foreign-key objects remain useful to the ORM for joins, but their target
    # master views are intentionally not part of this metadata. Preserve the
    # declared order instead of asking SQLAlchemy to resolve/sort those FKs.
    for table in metadata.tables.values():
        connection.execute(CreateTable(table, include_foreign_key_constraints=[]))
        for index in table.indexes:
            connection.execute(CreateIndex(index))



def _copy_public_master(connection, engine, key: str, table_name: str) -> None:
    """Copy formal reference rows to a private writable workspace table."""
    schema = _quote(engine, _assert_workspace_key(key))
    table = _quote(engine, table_name)
    private = f"{schema}.{table}"
    connection.execute(text(f"INSERT INTO {private} SELECT * FROM public.{table}"))
    # Do not lose private dishes previously created through the view overlay.
    if table_name == "kitchen_recipe" and inspect(connection).has_table(
        "practice_recipe_overlay", schema=key
    ):
        recipe_cols = [column.name for column in db.metadata.tables["kitchen_recipe"].columns]
        quoted_cols = ", ".join(_quote(engine, name) for name in recipe_cols)
        selected_cols = []
        for column_name in recipe_cols:
            if column_name == "name":
                # Preserve a private recipe even when the formal catalog later
                # introduced a dish with that name. Never discard its ID:
                # saved practice menu items may still point to that ID.
                selected_cols.append(
                    f"CASE WHEN EXISTS (SELECT 1 FROM {private} m "
                    "WHERE lower(m.name)=lower(p.name)) "
                    "THEN LEFT(p.name, 90) || '（練習原有' || p.id || '）' "
                    "ELSE p.name END"
                )
            else:
                selected_cols.append(f"p.{_quote(engine, column_name)}")
        connection.execute(text(
            f"INSERT INTO {private} ({quoted_cols}) "
            f"SELECT {', '.join(selected_cols)} "
            f"FROM {schema}.practice_recipe_overlay p "
            "ON CONFLICT (id) DO NOTHING"
        ))
    # The sequence must belong to the private table; never use public nextval.
    seq = connection.execute(text(
        "SELECT pg_get_serial_sequence(:table, 'id')"
    ), {"table": f"{key}.{table_name}"}).scalar()
    if seq:
        connection.execute(text(
            f"SELECT setval(CAST(:seq AS regclass), "
            f"COALESCE((SELECT MAX(id) FROM {private}), 1), "
            f"EXISTS(SELECT 1 FROM {private}))"
        ), {"seq": seq})


def _upgrade_postgres_master_views(connection, engine, key: str) -> None:
    """Upgrade legacy read-only views without touching saved practice orders."""
    schema = _quote(engine, _assert_workspace_key(key))
    connection.execute(
        text("SELECT pg_advisory_xact_lock(71320260, hashtext(:key))"),
        {"key": key},
    )
    for table_name in MASTER_TABLES:
        kind = connection.execute(text("""
            SELECT c.relkind FROM pg_class c
            JOIN pg_namespace n ON n.oid = c.relnamespace
            WHERE n.nspname=:schema AND c.relname=:table
        """), {"schema": key, "table": table_name}).scalar()
        if kind == "r":
            continue
        if kind != "v":
            raise PracticeAccountError("練習空間的主檔不存在，已停止操作以保護資料。")
        connection.execute(text(
            f"DROP VIEW {schema}.{_quote(engine, table_name)}"
        ))
        _clone_operational_tables(connection, key, (table_name,))
        _copy_public_master(connection, engine, key, table_name)


def _create_postgres_workspace(connection, engine, workspace_key: str) -> None:
    formal_target = (db.engine.url.host, db.engine.url.port, db.engine.url.database)
    practice_target = (engine.url.host, engine.url.port, engine.url.database)
    if formal_target != practice_target:
        raise PracticeAccountError(
            "練習區必須與正式主帳使用同一個 PostgreSQL 資料庫，"
            "才能安全建立獨立資料表。"
        )
    workspace_key = _assert_workspace_key(workspace_key)
    schema = _quote(engine, workspace_key)
    connection.execute(text(f"CREATE SCHEMA {schema}"))
    # Each account starts with its own fully writable reference data.
    _clone_operational_tables(connection, workspace_key, MASTER_TABLES)
    for table_name in MASTER_TABLES:
        _copy_public_master(connection, engine, workspace_key, table_name)
    _clone_operational_tables(connection, workspace_key)
    connection.execute(text(
        f"REVOKE ALL ON SCHEMA {schema} FROM PUBLIC, anon, authenticated"
    ))


def _sqlite_base_path(engine) -> Path:
    database = engine.url.database
    if not database or database == ":memory:":
        root = Path(current_app.instance_path)
        root.mkdir(parents=True, exist_ok=True)
        return root / "practice-control.db"
    return Path(database).resolve()


def _sqlite_workspace_path(engine, workspace_key: str) -> Path:
    workspace_key = _assert_workspace_key(workspace_key)
    base = _sqlite_base_path(engine)
    return base.with_name(f"{base.stem}-{workspace_key}{base.suffix or '.db'}")


def _sqlite_engine_cache():
    return current_app.extensions.setdefault("practice_workspace_engines", {})


def _sqlite_account_engine(base_engine, workspace_key: str):
    cache = _sqlite_engine_cache()
    if workspace_key not in cache:
        cache[workspace_key] = create_engine(
            f"sqlite:///{_sqlite_workspace_path(base_engine, workspace_key)}"
        )
    return cache[workspace_key]


def _create_sqlite_workspace(base_engine, workspace_key: str) -> None:
    target_engine = _sqlite_account_engine(base_engine, workspace_key)
    tables = [db.metadata.tables[name] for name in MASTER_TABLES + OPERATIONAL_TABLES]
    db.metadata.create_all(bind=target_engine, tables=tables)
    with db.engine.connect() as source, target_engine.begin() as target:
        for table_name in MASTER_TABLES:
            table = db.metadata.tables[table_name]
            rows = source.execute(select(table)).mappings().all()
            if rows:
                target.execute(table.insert(), [dict(row) for row in rows])


def _create_workspace(connection, engine, workspace_key: str) -> None:
    if _is_postgres(engine):
        _create_postgres_workspace(connection, engine, workspace_key)
    else:
        _create_sqlite_workspace(engine, workspace_key)


def _drop_workspace(connection, engine, workspace_key: str) -> None:
    workspace_key = _assert_workspace_key(workspace_key)
    if _is_postgres(engine):
        connection.execute(text(
            f"DROP SCHEMA IF EXISTS {_quote(engine, workspace_key)} CASCADE"
        ))
        # Execution-option engines share their underlying connection pool;
        # remove the wrapper, but never dispose the formal database pool.
        current_app.extensions.setdefault(
            "practice_postgres_workspace_engines", {}
        ).pop(workspace_key, None)
        return
    current_app.extensions.setdefault("practice_estimate_schema_checked", set()).discard(workspace_key)
    cached = _sqlite_engine_cache().pop(workspace_key, None)
    if cached is not None:
        cached.dispose()
    path = _sqlite_workspace_path(engine, workspace_key)
    if path.exists():
        path.unlink()


def cleanup_expired_accounts(*, now: datetime | None = None) -> int:
    engine = _practice_engine()
    _ensure_control(engine)
    now = now or _utcnow()
    ttl = int(current_app.config.get("PRACTICE_ACCOUNT_TTL_DAYS", 0))
    if ttl <= 0:
        return 0
    cutoff = now - timedelta(days=ttl)
    removed = 0
    with engine.begin() as connection:
        rows = connection.execute(text(
            f"SELECT id, workspace_key, last_active_at FROM {_control_name(engine)}"
        )).all()
        for row in rows:
            account = _row_dict(row)
            if account["last_active_at"] >= cutoff:
                continue
            _drop_workspace(connection, engine, account["workspace_key"])
            connection.execute(text(
                f"DELETE FROM {_control_name(engine)} WHERE id=:id"
            ), {"id": account["id"]})
            removed += 1
    return removed


def login_or_create_account(raw_name: str) -> tuple[dict, bool]:
    display_name, normalized_name = _normalize_name(raw_name)
    engine = _practice_engine()
    _ensure_control(engine)
    cleanup_expired_accounts()
    now = _utcnow()
    try:
        with engine.begin() as connection:
            if _is_postgres(engine):
                connection.execute(text("SELECT pg_advisory_xact_lock(71320260930)"))
            existing = _account_by(engine, "normalized_name", normalized_name, connection)
            if existing:
                connection.execute(text(
                    f"UPDATE {_control_name(engine)} SET last_active_at=:now WHERE id=:id"
                ), {"now": now, "id": existing["id"]})
                existing["last_active_at"] = now
                return existing, False
            count = connection.execute(text(
                f"SELECT COUNT(*) FROM {_control_name(engine)}"
            )).scalar_one()
            limit = int(current_app.config.get("PRACTICE_ACCOUNT_LIMIT", 20))
            if count >= limit:
                raise PracticeCapacityError(
                    f"練習帳號已達 {limit} 人上限，請聯絡管理者清理名額。"
                )
            account_id = uuid.uuid4().hex
            workspace_key = f"practice_{uuid.uuid4().hex[:24]}"
            _create_workspace(connection, engine, workspace_key)
            connection.execute(text(f"""
                INSERT INTO {_control_name(engine)}
                    (id, display_name, normalized_name, workspace_key, created_at, last_active_at)
                VALUES (:id, :display_name, :normalized_name, :workspace_key, :now, :now)
            """), {
                "id": account_id, "display_name": display_name,
                "normalized_name": normalized_name, "workspace_key": workspace_key,
                "now": now,
            })
            return {
                "id": account_id, "display_name": display_name,
                "normalized_name": normalized_name, "workspace_key": workspace_key,
                "created_at": now, "last_active_at": now,
            }, True
    except IntegrityError as exc:
        raise PracticeAccountError("帳號建立失敗，請稍後再試。") from exc


def get_account(account_id: str | None) -> dict | None:
    if not account_id:
        return None
    engine = _practice_engine()
    _ensure_control(engine)
    return _account_by(engine, "id", account_id)


def touch_current_account() -> dict | None:
    account = get_account(session.get("practice_account_id"))
    if account is None:
        return None
    now = _utcnow()
    ttl = int(current_app.config.get("PRACTICE_ACCOUNT_TTL_DAYS", 0))
    if ttl > 0 and account["last_active_at"] < now - timedelta(days=ttl):
        delete_account(account["id"])
        return None
    if now - account["last_active_at"] >= timedelta(minutes=30):
        engine = _practice_engine()
        with engine.begin() as connection:
            connection.execute(text(
                f"UPDATE {_control_name(engine)} SET last_active_at=:now WHERE id=:id"
            ), {"now": now, "id": account["id"]})
        account["last_active_at"] = now
    session["practice_account_name"] = account["display_name"]
    session["practice_workspace_key"] = account["workspace_key"]
    return account



def _ensure_postgres_recipe_overlay(connection, engine, workspace_key: str) -> None:
    """Let practice menus create private dishes while formal recipes remain read-only.

    The existing kitchen_recipe view continues to expose live formal recipes.
    Locally-created practice recipes use negative IDs and live only in the
    account's private table. An INSTEAD OF INSERT trigger handles ORM INSERT
    ... RETURNING without permitting UPDATE or DELETE on the shared view.
    """
    schema = _quote(engine, _assert_workspace_key(workspace_key))
    if inspect(connection).has_table("practice_recipe_overlay", schema=workspace_key):
        return

    # LIKE tracks the formal catalog's columns, including any additive fields.
    # Never copy its SERIAL default: that would consume the *formal* ID sequence.
    connection.execute(text(f"""
        CREATE TABLE {schema}.practice_recipe_overlay (
            LIKE public.kitchen_recipe INCLUDING CONSTRAINTS,
            PRIMARY KEY (id), UNIQUE (name)
        )
    """))
    connection.execute(text(f"""
        CREATE SEQUENCE {schema}.practice_recipe_id_seq
        AS integer INCREMENT BY -1 MINVALUE -2147483648
        MAXVALUE -1 START WITH -1
    """))
    # CREATE OR REPLACE retains the original view's column names/types, so
    # existing ORM relationships and workspace menu items still work.
    connection.execute(text(f"""
        CREATE OR REPLACE VIEW {schema}.kitchen_recipe
        WITH (security_invoker=true) AS
        SELECT * FROM public.kitchen_recipe
        UNION ALL
        SELECT * FROM {schema}.practice_recipe_overlay
    """))
    connection.execute(text(f"""
        CREATE FUNCTION {schema}.insert_practice_recipe()
        RETURNS trigger LANGUAGE plpgsql AS $practice_insert$
        BEGIN
            IF NEW.id IS NOT NULL THEN
                RAISE EXCEPTION 'Practice recipe IDs must be generated';
            END IF;
            IF EXISTS (SELECT 1 FROM public.kitchen_recipe WHERE name = NEW.name) THEN
                RAISE EXCEPTION 'Recipe already exists in the formal catalog';
            END IF;
            NEW.id := nextval('{workspace_key}.practice_recipe_id_seq');
            NEW.active := COALESCE(NEW.active, FALSE);
            NEW.created_at := COALESCE(NEW.created_at, NOW());
            NEW.updated_at := COALESCE(NEW.updated_at, NOW());
            INSERT INTO {schema}.practice_recipe_overlay SELECT NEW.*;
            RETURN NEW;
        END
        $practice_insert$;
    """))
    connection.execute(text(f"""
        CREATE TRIGGER insert_practice_recipe
        INSTEAD OF INSERT ON {schema}.kitchen_recipe
        FOR EACH ROW EXECUTE FUNCTION {schema}.insert_practice_recipe()
    """))
    connection.execute(text(
        f"REVOKE ALL ON FUNCTION {schema}.insert_practice_recipe() FROM PUBLIC"
    ))


def _ensure_workspace_estimate_columns(account: dict) -> None:
    """Idempotently upgrade operational columns in an existing named workspace."""
    engine = _practice_engine()
    key = _assert_workspace_key(account["workspace_key"])
    cache = current_app.extensions.setdefault("practice_estimate_schema_checked", set())
    if key in cache:
        return
    is_postgres = _is_postgres(engine)
    bind = engine if is_postgres else _sqlite_account_engine(engine, key)
    with bind.begin() as connection:
        if is_postgres:
            _upgrade_postgres_master_views(connection, engine, key)
        # Old practice tables do not automatically receive newly added
        # columns from public; missing ones cause UndefinedColumn on SELECT.
        upgrades = {
            "kitchen_purchase_order_item": {
                "dish_estimates_json": "TEXT NOT NULL DEFAULT '{}'",
                "dish_adjustment_qty": "NUMERIC(16, 4)",
            },
            "kitchen_menu_plan_item": {
                "ingredients_override_json": "TEXT",
            },
            "kitchen_daily_dish_note": {
                "count_override_fields": "TEXT",
            },
        }
        for table_name, additions in upgrades.items():
            cols = {c["name"] for c in inspect(connection).get_columns(
                table_name, schema=key if is_postgres else None,
            )}
            qualified = (
                f"{_quote(engine, key)}.{_quote(engine, table_name)}"
                if is_postgres else _quote(engine, table_name)
            )
            for column_name, definition in additions.items():
                if column_name not in cols:
                    connection.execute(text(
                        f"ALTER TABLE {qualified} ADD COLUMN "
                        f"{_quote(engine, column_name)} {definition}"
                    ))
    cache.add(key)


def workspace_engine(account: dict | None = None):
    account = account or get_account(session.get("practice_account_id"))
    if account is None:
        raise PracticeAccountError("練習帳號已不存在。")
    engine = _practice_engine()
    _ensure_workspace_estimate_columns(account)
    if _is_postgres(engine):
        schema = _assert_workspace_key(account["workspace_key"])
        # Keep the *same* OptionEngine identity throughout the request.
        # SQLAlchemy Session keys transactional connections by bind identity:
        # constructing a fresh execution_options() wrapper for each get_bind()
        # creates independent connections. PostgreSQL reads then cannot see
        # uncommitted menu plans / recipes flushed on a different connection.
        cache = current_app.extensions.setdefault("practice_postgres_workspace_engines", {})
        if schema not in cache:
            cache[schema] = engine.execution_options(
                schema_translate_map={None: schema}
            )
        return cache[schema]
    return _sqlite_account_engine(engine, account["workspace_key"])


def list_accounts() -> list[dict]:
    cleanup_expired_accounts()
    engine = _practice_engine()
    with engine.connect() as connection:
        rows = connection.execute(text(
            f"SELECT id, display_name, normalized_name, workspace_key, created_at, last_active_at "
            f"FROM {_control_name(engine)} ORDER BY last_active_at DESC, display_name"
        )).all()
    accounts = [_row_dict(row) for row in rows]
    now = _utcnow()
    ttl = int(current_app.config.get("PRACTICE_ACCOUNT_TTL_DAYS", 0))
    for account in accounts:
        if ttl > 0:
            remaining = timedelta(days=ttl) - (now - account["last_active_at"])
            account["days_remaining"] = max(0, remaining.days + (1 if remaining.seconds else 0))
        else:
            account["days_remaining"] = None
        try:
            account_engine = workspace_engine(account)
            with account_engine.connect() as connection:
                account["plan_count"] = connection.execute(
                    select(func.count()).select_from(db.metadata.tables["kitchen_menu_plan"])
                ).scalar_one()
                account["order_count"] = connection.execute(
                    select(func.count()).select_from(db.metadata.tables["kitchen_purchase_order"])
                ).scalar_one()
        except Exception:
            account["plan_count"] = None
            account["order_count"] = None
    return accounts


def delete_account(account_id: str) -> bool:
    engine = _practice_engine()
    _ensure_control(engine)
    with engine.begin() as connection:
        account = _account_by(engine, "id", account_id, connection)
        if account is None:
            return False
        _drop_workspace(connection, engine, account["workspace_key"])
        connection.execute(text(
            f"DELETE FROM {_control_name(engine)} WHERE id=:id"
        ), {"id": account_id})
    return True


def reset_account(account_id: str) -> bool:
    engine = _practice_engine()
    _ensure_control(engine)
    with engine.begin() as connection:
        account = _account_by(engine, "id", account_id, connection)
        if account is None:
            return False
        _drop_workspace(connection, engine, account["workspace_key"])
        _create_workspace(connection, engine, account["workspace_key"])
        connection.execute(text(
            f"UPDATE {_control_name(engine)} SET last_active_at=:now WHERE id=:id"
        ), {"now": _utcnow(), "id": account_id})
    return True


def ensure_practice_database() -> None:
    """Compatibility entry point: initialize account control storage only."""
    _ensure_control(_practice_engine())


def activate_account(account: dict) -> None:
    db.session.remove()
    session.pop("practice_selected_week", None)
    session["kitchen_practice"] = True
    session["practice_account_id"] = account["id"]
    session["practice_account_name"] = account["display_name"]
    session["practice_workspace_key"] = account["workspace_key"]


def clear_practice_session() -> None:
    db.session.remove()
    for key in (
        "kitchen_practice", "practice_account_id", "practice_account_name",
        "practice_workspace_key", "practice_selected_week",
    ):
        session.pop(key, None)


def practice_master_write_blocked(endpoint: str | None) -> bool:
    return endpoint in {
        "order_tool.schools", "order_tool.school_update", "order_tool.school_toggle",
        "order_tool.suppliers", "order_tool.supplier_update", "order_tool.supplier_toggle",
        "order_tool.supplier_item_create", "order_tool.supplier_item_update",
        "order_tool.supplier_item_delete", "order_tool.ingredients",
        "order_tool.ingredient_update", "order_tool.ingredient_toggle",
        "order_tool.ingredient_nutrition_update", "order_tool.recipes",
        "order_tool.recipe_update", "order_tool.recipe_toggle", "order_tool.recipe_delete",
        "order_tool.recipe_category_update", "order_tool.recipe_copy",
        "order_tool.recipe_ingredient_add", "order_tool.recipe_ingredient_update",
        "order_tool.recipe_ingredient_delete", "order_tool.recipe_tags_save",
        "order_tool.ai_menu_swap_new", "order_tool.catalog_import",
        "order_tool.catalog_import_apply", "order_tool.procurement_manual_item_create",
    }
