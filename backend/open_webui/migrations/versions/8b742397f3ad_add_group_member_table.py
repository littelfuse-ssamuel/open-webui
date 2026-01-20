"""Add group_member table

Revision ID: 8b742397f3ad
Revises: b9f03403a733
Create Date: 2025-01-20 22:00:00.000000

This migration adds the group_member table to replace the user_ids JSON column
in the group table. This is required for compatibility with Open WebUI 0.6.37+.

"""

from typing import Sequence, Union
from alembic import op
import sqlalchemy as sa
import json
import uuid

# revision identifiers, used by Alembic.
revision: str = "8b742397f3ad"  # Generate with: python -c "import uuid; print(uuid.uuid4().hex[:12])"
down_revision: Union[str, None] = "b9f03403a733"
branch_labels: Union[str, Sequence[str], None] = None
depends_on: Union[str, Sequence[str], None] = None


def upgrade() -> None:
    conn = op.get_bind()
    dialect = conn.dialect.name
    
    # 1. Create the group_member table
    op.create_table(
        "group_member",
        sa.Column("id", sa.Text(), primary_key=True),
        sa.Column("group_id", sa.Text(), nullable=False),
        sa.Column("user_id", sa.Text(), nullable=False),
        sa.Column("created_at", sa.BigInteger(), nullable=True),
    )
    
    # Create indexes
    op.create_index("ix_group_member_group_id", "group_member", ["group_id"])
    op.create_index("ix_group_member_user_id", "group_member", ["user_id"])
    
    # Create unique constraint on (group_id, user_id)
    op.create_unique_constraint(
        "uq_group_member_group_user",
        "group_member",
        ["group_id", "user_id"]
    )
    
    # Add foreign keys (with CASCADE delete)
    if dialect == "postgresql":
        op.create_foreign_key(
            "fk_group_member_group_id",
            "group_member",
            "group",
            ["group_id"],
            ["id"],
            ondelete="CASCADE"
        )
        op.create_foreign_key(
            "fk_group_member_user_id",
            "group_member",
            "user",
            ["user_id"],
            ["id"],
            ondelete="CASCADE"
        )
    
    # 2. Migrate data from group.user_ids JSON to group_member table
    import time
    now = int(time.time())
    
    # Get all groups with user_ids
    if dialect == "postgresql":
        # PostgreSQL: user_ids is JSON type
        results = conn.execute(sa.text('''
            SELECT id, user_ids FROM "group" WHERE user_ids IS NOT NULL
        ''')).fetchall()
    else:
        # SQLite
        results = conn.execute(sa.text('''
            SELECT id, user_ids FROM "group" WHERE user_ids IS NOT NULL
        ''')).fetchall()
    
    for row in results:
        group_id = row[0]
        user_ids_raw = row[1]
        
        # Parse user_ids (might be JSON string or already parsed)
        if user_ids_raw is None:
            continue
            
        if isinstance(user_ids_raw, str):
            try:
                user_ids = json.loads(user_ids_raw)
            except (json.JSONDecodeError, TypeError):
                continue
        elif isinstance(user_ids_raw, list):
            user_ids = user_ids_raw
        else:
            continue
        
        if not isinstance(user_ids, list):
            continue
            
        for user_id in user_ids:
            if user_id:
                member_id = str(uuid.uuid4())
                try:
                    conn.execute(
                        sa.text('''
                            INSERT INTO group_member (id, group_id, user_id, created_at)
                            VALUES (:id, :group_id, :user_id, :created_at)
                            ON CONFLICT (group_id, user_id) DO NOTHING
                        '''),
                        {
                            "id": member_id,
                            "group_id": group_id,
                            "user_id": user_id,
                            "created_at": now
                        }
                    )
                except Exception as e:
                    # Skip duplicates or invalid foreign keys
                    print(f"Skipping group_member insert for group={group_id}, user={user_id}: {e}")
                    continue
    
    # 3. Drop the user_ids column from group table
    # Note: We keep the column for now in case rollback is needed
    # You can uncomment this after verifying the migration works:
    # op.drop_column("group", "user_ids")


def downgrade() -> None:
    conn = op.get_bind()
    dialect = conn.dialect.name
    
    # 1. Ensure user_ids column exists (add it back if it was dropped)
    try:
        op.add_column("group", sa.Column("user_ids", sa.JSON(), nullable=True))
    except Exception:
        pass  # Column might already exist
    
    # 2. Migrate data back from group_member to group.user_ids
    if dialect == "postgresql":
        conn.execute(sa.text('''
            UPDATE "group" g
            SET user_ids = (
                SELECT json_agg(gm.user_id)
                FROM group_member gm
                WHERE gm.group_id = g.id
            )
        '''))
    else:
        # SQLite - need to do this row by row
        groups = conn.execute(sa.text('SELECT id FROM "group"')).fetchall()
        for (group_id,) in groups:
            members = conn.execute(
                sa.text('SELECT user_id FROM group_member WHERE group_id = :gid'),
                {"gid": group_id}
            ).fetchall()
            user_ids = [m[0] for m in members]
            conn.execute(
                sa.text('UPDATE "group" SET user_ids = :uids WHERE id = :gid'),
                {"uids": json.dumps(user_ids), "gid": group_id}
            )
    
    # 3. Drop foreign keys if PostgreSQL
    if dialect == "postgresql":
        try:
            op.drop_constraint("fk_group_member_user_id", "group_member", type_="foreignkey")
            op.drop_constraint("fk_group_member_group_id", "group_member", type_="foreignkey")
        except Exception:
            pass
    
    # 4. Drop indexes and table
    op.drop_index("ix_group_member_user_id", table_name="group_member")
    op.drop_index("ix_group_member_group_id", table_name="group_member")
    op.drop_constraint("uq_group_member_group_user", "group_member", type_="unique")
    op.drop_table("group_member")