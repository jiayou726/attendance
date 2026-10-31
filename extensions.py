from flask import has_request_context, request, session
from flask_migrate import Migrate
from flask_sqlalchemy import SQLAlchemy
from flask_sqlalchemy.session import Session as FlaskSQLAlchemySession


class KitchenRoutingSession(FlaskSQLAlchemySession):
    """Route kitchen practice requests to an isolated database bind.

    A dedicated practice-workspace URL scope is required, alongside the
    signed account session. Formal URLs are *always* served by public data,
    even if this browser also has an active practice account.
    """

    def get_bind(self, mapper=None, clause=None, bind=None, **kwargs):
        if (
            bind is None
            and has_request_context()
            and session.get("kitchen_practice")
            and request.path.startswith("/admin/order-tool/practice-workspace/")
        ):
            # Import lazily to avoid the extensions -> models -> extensions
            # cycle during application startup.
            from services.practice_database import workspace_engine
            return workspace_engine()
        return super().get_bind(mapper=mapper, clause=clause, bind=bind, **kwargs)


db = SQLAlchemy(session_options={"class_": KitchenRoutingSession})
migrate = Migrate()
