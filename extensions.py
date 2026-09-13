from flask import has_request_context, request, session
from flask_migrate import Migrate
from flask_sqlalchemy import SQLAlchemy
from flask_sqlalchemy.session import Session as FlaskSQLAlchemySession


class KitchenRoutingSession(FlaskSQLAlchemySession):
    """Route kitchen practice requests to an isolated database bind.

    The rest of the application keeps using the normal database.  Once the
    practice flag is present in the browser session, every ORM read/write
    under /admin/order-tool is transparently sent to the ``practice`` bind.
    This keeps existing kitchen routes fully functional without sprinkling
    practice checks through every CRUD handler.
    """

    def get_bind(self, mapper=None, clause=None, bind=None, **kwargs):
        if (
            bind is None
            and has_request_context()
            and session.get("kitchen_practice")
            and request.path.startswith("/admin/order-tool")
        ):
            practice_engine = self._db.engines.get("practice")
            if practice_engine is None:
                raise RuntimeError("Practice database bind is not configured.")
            return practice_engine
        return super().get_bind(mapper=mapper, clause=clause, bind=bind, **kwargs)


db = SQLAlchemy(session_options={"class_": KitchenRoutingSession})
migrate = Migrate()
