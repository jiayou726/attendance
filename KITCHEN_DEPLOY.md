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

## 3. Supabase / migration warning

This repository has historical Alembic migration drift around old attendance table names. Therefore **do not blindly run `flask db migrate` against production Supabase**.

The manually reviewed migration is:

`migrations/versions/kitchen20260813_prod.py`

It is intentionally written to touch only these tables:

- `kitchen_school`
- `kitchen_supplier`
- `kitchen_ingredient`
- `kitchen_recipe`
- `kitchen_recipe_ingredient`
- `kitchen_menu_plan`
- `kitchen_menu_plan_item`
- `kitchen_menu_assignment`
- `kitchen_purchase_order`
- `kitchen_purchase_order_item`

It must never drop, rename or alter `employee` / `checkin`.

### Recommended release procedure

1. Clone the current production database to a staging/test PostgreSQL project when possible.
2. Inspect `alembic_version` and current table names first.
3. Confirm the staging database already has the same attendance schema used by production.
4. Run the reviewed migration on staging.
5. Re-run the kitchen regression tests against staging.
6. Verify attendance/punch still works.
7. Only then apply the exact same reviewed migration to production.

If the production database has no `alembic_version` table, do not guess or stamp it without first inspecting the schema. Ask for a schema dump / table list and decide the correct stamp point before running Alembic.

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
