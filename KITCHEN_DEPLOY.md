# Kitchen v2 deployment / verification

`kitchen-v2` is designed to keep the existing attendance models (`employee`, `checkin`) untouched and add the group-meal workflow under `kitchen_*` tables.

## 1. Before production

Do not merge or deploy until the branch passes:

```bash
python -m pip install -r requirements-dev.txt
python -m compileall -q app.py config.py models.py blueprints tests
pytest -q tests/test_kitchen.py
```

Browser verification should cover at least 375×812, 430×932, 768×1024 and 1440×900.

## 2. Required production environment variables

Use the deployment platform's secret/environment settings. Do not put real values in `.env.example` or GitHub.

Required:

- `DATABASE_URL`: Supabase PostgreSQL server connection string used by Flask/SQLAlchemy.
- `SECRET_KEY`: long random secret.
- `ADMIN_HR_PASSWORD` and/or `ADMIN_MGR_PASSWORD`: strong admin password(s).
- `PRODUCTION=1`
- `AUTO_CREATE_DB=0`
- `KITCHEN_CSRF_ENABLED=1`

When `PRODUCTION=1`, the app intentionally refuses to start if `SECRET_KEY` or all admin passwords are missing, or if kitchen CSRF is disabled.

## 3. Supabase / schema warning

This repository has historical Alembic migration drift around old attendance table names. Therefore **do not blindly run `flask db migrate` or a full historical `flask db upgrade` against production Supabase**.

Two kitchen-only tools are provided:

- `scripts/bootstrap_kitchen_schema.py`
- `migrations/versions/kitchen20260813_prod.py`

Both are intentionally limited to `kitchen_*` tables and must never drop, rename or alter `employee` / `checkin`.

### Case A: production has no kitchen_* tables yet

This is the safest initial deployment path.

After setting the real server-side `DATABASE_URL` and production secrets, run:

```bash
python scripts/bootstrap_kitchen_schema.py
```

The script creates only the final kitchen tables in dependency order. It does not call `db.create_all()` for attendance tables.

If it finds an existing kitchen table whose columns do not match the final model, it **refuses to alter it** and exits. This is intentional fail-closed behavior.

### Case B: production/staging already has older MVP kitchen_* tables

Do not use the bootstrap script to silently patch them. Use a staging clone first and review/apply:

`migrations/versions/kitchen20260813_prod.py`

The reviewed migration can add the missing production columns/tables while remaining kitchen-only.

### Recommended release procedure

1. Inspect the current Supabase table list and `alembic_version` first.
2. Prefer a staging/test Supabase project cloned from production.
3. Confirm staging attendance tables match production.
4. If no kitchen tables exist, test `python scripts/bootstrap_kitchen_schema.py` on staging.
5. If older kitchen tables exist, test the reviewed kitchen-only migration on staging instead.
6. Run the full kitchen regression tests.
7. Verify existing attendance/admin pages and `/punch` still work.
8. Only then repeat the same kitchen-only schema operation on production.

If production has no `alembic_version` table, do not guess or stamp it without inspecting the schema first.

## 4. Business-flow acceptance test

Use this exact test case:

- School: 內小
- Headcount: 801
- Supplier: 測試肉品
- Ingredient: 骨腿丁
- Purchase unit: kg
- `grams_per_purchase_unit = 1000`
- Price: 82 / kg
- Recipe: 南洋綠咖哩雞
- Serving output: 95 g/person
- AP ingredient amount: 88 g/person

Expected demand:

`88 × 801 = 70,488 g = 70.488 kg`

Expected raw-material cost at 82/kg:

`70.488 × 82 = 5,780.016`

The 95 g serving-output field must not be used to calculate bone-in chicken procurement.

Then add a second school with 586 diners. Expected chicken demand:

`(801 + 586) × 88 = 122,056 g = 122.056 kg`

## 5. Purchase snapshot acceptance test

1. Generate a draft purchase order.
2. Confirm the order while chicken is 82/kg and the recipe uses 88 g/person.
3. Change the ingredient master price to 90/kg.
4. Change the recipe AP amount to 90 g/person.
5. Open the already-confirmed historical order.

The confirmed order must still show its original required grams, original 82 price snapshot and original amount. It must not be recalculated from current master data.

## 6. Mobile / desktop acceptance

Kitchen templates load `static/kitchen_mobile.css` and `static/kitchen_ui.js` directly. The old response-injection approach has been removed.

Check:

- phone bottom navigation and safe-area spacing;
- 44–48 px touch targets;
- iPhone inputs use 16 px text to prevent Safari auto-zoom;
- wide tables scroll horizontally on phones;
- desktop retains multi-column cards and normal tables;
- print view hides navigation;
- attendance pages do not load kitchen CSS/JS.

## 7. Security notes

- Old hard-coded admin passwords were removed from current branch code. If any old password was used in real life, treat it as compromised and replace it.
- `SECRET_KEY` is no longer a fixed public string.
- `/admin/*` is login-protected; `/punch` remains outside that guard.
- Kitchen mutation POSTs use a session CSRF token.
- `.env` and local SQLite files are ignored for future commits.
- The repository history may still contain previously committed secrets or databases; `.gitignore` does not erase Git history. Review history separately before treating the repository as clean for sensitive production use.

## 8. Do not merge automatically

Keep work on `kitchen-v2` until automated tests plus browser tests pass. Merge to `main` should be an explicit final decision after review.
