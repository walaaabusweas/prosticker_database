import os
import uuid
from datetime import datetime, timezone
from sqlalchemy import create_engine, String, Text, Boolean, DateTime, Integer, Float, UniqueConstraint, select, func
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column, sessionmaker
from sqlalchemy.types import JSON
from sqlalchemy.dialects.postgresql import JSONB

JSONType = JSON().with_variant(JSONB, 'postgresql')


def utcnow():
    return datetime.now(timezone.utc)


class Base(DeclarativeBase):
    pass


class SchemaType(Base):
    __tablename__ = 'schema_types'
    final_type: Mapped[str] = mapped_column(String(80), primary_key=True)
    category: Mapped[str | None] = mapped_column(String(80))
    arabic_name: Mapped[str | None] = mapped_column(String(200))
    subtypes_text: Mapped[str | None] = mapped_column(Text)
    status: Mapped[str | None] = mapped_column(String(50))
    products_covered: Mapped[str | None] = mapped_column(Text)


class SchemaSubtype(Base):
    __tablename__ = 'schema_subtypes'
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    final_type: Mapped[str] = mapped_column(String(80), index=True)
    subtype: Mapped[str] = mapped_column(String(100))
    arabic_name: Mapped[str | None] = mapped_column(String(200))
    kind: Mapped[str | None] = mapped_column(String(80))
    description: Mapped[str | None] = mapped_column(Text)
    __table_args__ = (UniqueConstraint('final_type', 'subtype', name='uq_schema_subtype'),)


class PropertyDefinition(Base):
    __tablename__ = 'property_definitions'
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    final_type: Mapped[str] = mapped_column(String(80), index=True)
    subtype: Mapped[str | None] = mapped_column(String(100), index=True)
    scope: Mapped[str] = mapped_column(String(40))
    property: Mapped[str] = mapped_column(String(120), index=True)
    description_ar: Mapped[str | None] = mapped_column(String(300))
    data_type: Mapped[str] = mapped_column(String(50))
    allowed_values: Mapped[str | None] = mapped_column(Text)
    required: Mapped[bool] = mapped_column(Boolean, default=False)
    notes: Mapped[str | None] = mapped_column(Text)


class LookupValue(Base):
    __tablename__ = 'lookup_values'
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    lookup_type: Mapped[str] = mapped_column(String(60), index=True)
    value: Mapped[str] = mapped_column(String(100))
    english: Mapped[str | None] = mapped_column(String(200))
    arabic: Mapped[str | None] = mapped_column(String(200))
    role: Mapped[str | None] = mapped_column(String(200))
    notes: Mapped[str | None] = mapped_column(Text)
    __table_args__ = (UniqueConstraint('lookup_type', 'value', name='uq_lookup_value'),)


class Product(Base):
    __tablename__ = 'products'
    id: Mapped[str] = mapped_column(String(36), primary_key=True, default=lambda: str(uuid.uuid4()))
    product_name: Mapped[str] = mapped_column(String(300), index=True)
    product_type: Mapped[str] = mapped_column(String(80), index=True)
    subtype: Mapped[str | None] = mapped_column(String(100), index=True)
    description: Mapped[str | None] = mapped_column(Text)
    attributes: Mapped[dict] = mapped_column(JSONType, default=dict)
    source: Mapped[str | None] = mapped_column(String(100), default='client_entry')
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow)
    updated_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=utcnow, onupdate=utcnow)


DATABASE_URL = os.getenv('DATABASE_URL', 'sqlite:///./data/printing_company.db')
connect_args = {'check_same_thread': False} if DATABASE_URL.startswith('sqlite') else {}
engine = create_engine(DATABASE_URL, future=True, pool_pre_ping=True, connect_args=connect_args)
SessionLocal = sessionmaker(bind=engine, autoflush=False, expire_on_commit=False)


def init_db():
    Base.metadata.create_all(engine)


def seed_schema(schema: dict):
    init_db()
    with SessionLocal() as db:
        if db.scalar(select(func.count()).select_from(SchemaType)) == 0:
            for r in schema['final_types']:
                db.add(SchemaType(
                    final_type=r.get('Final Type'), category=r.get('Category'),
                    arabic_name=r.get('الاسم بالعربي'), subtypes_text=r.get('Subtypes / Variants'),
                    status=r.get('Status'), products_covered=r.get('Products covered / examples')
                ))
            for r in schema['subtypes']:
                db.add(SchemaSubtype(final_type=r.get('Final Type'), subtype=r.get('Subtype'),
                                     arabic_name=r.get('الاسم بالعربي'), kind=r.get('Kind'), description=r.get('الوصف')))
            for r in schema['properties']:
                db.add(PropertyDefinition(**r))
            for lookup_type, rows in schema['lookups'].items():
                for r in rows:
                    db.add(LookupValue(
                        lookup_type=lookup_type, value=r.get('Value'), english=r.get('English'),
                        arabic=r.get('العربي'), role=r.get('Role'),
                        notes=r.get('Example / Notes') or r.get('Meaning') or r.get('Example')
                    ))
            db.commit()


def list_products(search: str = ''):
    with SessionLocal() as db:
        q = select(Product).order_by(Product.updated_at.desc())
        if search.strip():
            q = q.where(Product.product_name.ilike(f'%{search.strip()}%'))
        return list(db.scalars(q).all())


def get_product(product_id: str):
    with SessionLocal() as db:
        return db.get(Product, product_id)


def save_product(data: dict, product_id: str | None = None):
    with SessionLocal() as db:
        product = db.get(Product, product_id) if product_id else None
        if product is None:
            product = Product(id=str(uuid.uuid4()), source='client_entry')
            db.add(product)
        for k, v in data.items():
            if hasattr(product, k):
                setattr(product, k, v)
        product.updated_at = utcnow()
        db.commit()
        db.refresh(product)
        return product


def delete_product(product_id: str):
    with SessionLocal() as db:
        product = db.get(Product, product_id)
        if product:
            db.delete(product); db.commit()


def counts():
    with SessionLocal() as db:
        return {
            'products': db.scalar(select(func.count()).select_from(Product)) or 0,
            'types': db.scalar(select(func.count()).select_from(SchemaType)) or 0,
            'subtypes': db.scalar(select(func.count()).select_from(SchemaSubtype)) or 0,
        }
