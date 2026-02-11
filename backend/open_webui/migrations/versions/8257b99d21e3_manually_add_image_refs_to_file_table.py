"""Manually add image_refs to file table

Revision ID: 8257b99d21e3
Revises: c440947495f3
Create Date: 2025-08-28 16:30:00.000000

"""
from typing import Sequence, Union

from alembic import op
import sqlalchemy as sa
from sqlalchemy import inspect


# revision identifiers, used by Alembic.
revision: str = "8257b99d21e3"
down_revision: Union[str, None] = "c440947495f3"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def _has_image_refs_column() -> bool:
    bind = op.get_bind()
    inspector = inspect(bind)
    columns = {col["name"] for col in inspector.get_columns("file")}
    return "image_refs" in columns


def upgrade() -> None:
    if _has_image_refs_column():
        return

    # Use batch mode for SQLite compatibility
    with op.batch_alter_table("file", schema=None) as batch_op:
        batch_op.add_column(sa.Column("image_refs", sa.JSON(), nullable=True))


def downgrade() -> None:
    if not _has_image_refs_column():
        return

    # Use batch mode for SQLite compatibility
    with op.batch_alter_table("file", schema=None) as batch_op:
        batch_op.drop_column("image_refs")
