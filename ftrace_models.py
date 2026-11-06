"""Shared ftrace records, isolated from menus and procurement."""
from datetime import datetime, timezone
from extensions import db

def now():
    return datetime.now(timezone.utc)

class FtraceDocument(db.Model):
    __tablename__ = 'kitchen_ftrace_document'
    id = db.Column(db.Integer, primary_key=True)
    day = db.Column(db.Date, nullable=False)
    kind = db.Column(db.String(20), nullable=False)
    version = db.Column(db.Integer, nullable=False, default=1)
    status = db.Column(db.String(20), nullable=False, default='draft')
    payload_json = db.Column(db.Text, nullable=False)
    updated_at = db.Column(db.DateTime(timezone=True), nullable=False, default=now)
    __table_args__ = (db.UniqueConstraint('day','kind',name='uq_ftrace_document_day_kind'),)

class FtraceRevision(db.Model):
    __tablename__ = 'kitchen_ftrace_revision'
    id = db.Column(db.Integer, primary_key=True)
    document_id = db.Column(db.Integer, nullable=False, index=True)
    version = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), nullable=False)
    payload_json = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), nullable=False, default=now)
    __table_args__ = (db.UniqueConstraint('document_id','version',name='uq_ftrace_revision_version'),)

class FtraceImport(db.Model):
    __tablename__ = 'kitchen_ftrace_import'
    id = db.Column(db.Integer, primary_key=True)
    digest = db.Column(db.String(64), nullable=False, unique=True)
    filename = db.Column(db.String(255), nullable=False)
    original_file = db.Column(db.LargeBinary, nullable=False)
    summary_json = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime(timezone=True), nullable=False, default=now)

class FtraceReceipt(db.Model):
    __tablename__ = 'kitchen_ftrace_receipt'
    id = db.Column(db.Integer, primary_key=True)
    import_id = db.Column(db.Integer, nullable=False, index=True)
    source_row = db.Column(db.Integer, nullable=False)
    report_date = db.Column(db.String(10), nullable=False, index=True)
    material_code = db.Column(db.String(40), nullable=False, index=True)
    name = db.Column(db.String(255), nullable=False)
    payload_json = db.Column(db.Text, nullable=False)
    review_note = db.Column(db.Text, nullable=False, default='')
    disabled = db.Column(db.Boolean, nullable=False, default=False)
    version = db.Column(db.Integer, nullable=False, default=1)

TABLES = (FtraceDocument.__table__, FtraceRevision.__table__, FtraceImport.__table__, FtraceReceipt.__table__)

def ensure_schema(engine):
    from sqlalchemy import text
    with engine.begin() as connection:
        if engine.dialect.name == 'postgresql':
            connection.execute(text('SELECT pg_advisory_xact_lock(71320260, 6)'))
        db.metadata.create_all(connection, tables=TABLES)
        if engine.dialect.name == 'postgresql':
            for table in TABLES:
                connection.execute(text(f'ALTER TABLE public.{table.name} ENABLE ROW LEVEL SECURITY'))
                connection.execute(text(f'REVOKE ALL ON TABLE public.{table.name} FROM PUBLIC, anon, authenticated'))
