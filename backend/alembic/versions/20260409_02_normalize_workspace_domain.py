"""normalize workspace domain

Revision ID: 20260409_02
Revises: 20260409_01
Create Date: 2026-04-09 00:30:00.000000

"""

from __future__ import annotations

import json
import re
from datetime import UTC, datetime
from uuid import uuid4

from alembic import op
import sqlalchemy as sa


# revision identifiers, used by Alembic.
revision = "20260409_02"
down_revision = "20260409_01"
branch_labels = None
depends_on = None

SYSTEM_USER_ID = "user_system"
SYSTEM_USER_EMAIL = "system@startsheet.local"
SYSTEM_USER_NAME = "Startsheet System"

DEFAULT_VIEW_CONFIG = {
    "columnOrder": [],
    "columnVisibility": {},
    "sortModel": [],
    "filterModel": {},
    "grouping": [],
    "layout": {"rowHeightMode": "comfortable"},
}


def _timestamp() -> datetime:
    return datetime.now(UTC)


def _slugify(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", (value or "").strip().lower()).strip("-")
    return slug or "item"


def _unique_slug(value: str, existing_slugs: set[str]) -> str:
    base_slug = _slugify(value)
    candidate = base_slug
    suffix = 2

    while candidate in existing_slugs:
        candidate = f"{base_slug}-{suffix}"
        suffix += 1

    return candidate


def _unique_key(value: str, existing_keys: set[str]) -> str:
    base_key = _slugify(value).replace("-", "_")
    candidate = base_key
    suffix = 2

    while candidate in existing_keys:
        candidate = f"{base_key}_{suffix}"
        suffix += 1

    return candidate


def _default_view_config() -> dict[str, object]:
    return {
        "columnOrder": [],
        "columnVisibility": {},
        "sortModel": [],
        "filterModel": {},
        "grouping": [],
        "layout": {"rowHeightMode": "comfortable"},
    }


def _sheet_kind_from_icon(icon: str | None) -> str:
    return "data" if not icon or icon == "grid" else icon


def upgrade() -> None:
    connection = op.get_bind()
    now = _timestamp()

    op.create_table(
        "users",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("email", sa.String(length=255), nullable=False),
        sa.Column("full_name", sa.String(length=120), nullable=False),
        sa.Column("avatar_url", sa.String(length=512), nullable=True),
        sa.Column("is_active", sa.Boolean(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_users_email", "users", ["email"], unique=True)

    connection.execute(
        sa.text(
            """
            INSERT INTO users (id, email, full_name, avatar_url, is_active, created_at, updated_at)
            VALUES (:id, :email, :full_name, :avatar_url, :is_active, :created_at, :updated_at)
            """
        ),
        {
            "id": SYSTEM_USER_ID,
            "email": SYSTEM_USER_EMAIL,
            "full_name": SYSTEM_USER_NAME,
            "avatar_url": None,
            "is_active": True,
            "created_at": now,
            "updated_at": now,
        },
    )

    op.add_column("workspaces", sa.Column("slug", sa.String(length=120), nullable=True))
    op.add_column("workspaces", sa.Column("owner_id", sa.String(length=32), nullable=True))

    workspace_rows = connection.execute(
        sa.text("SELECT id, name, position, created_at, updated_at FROM workspaces ORDER BY position, created_at, id")
    ).mappings().all()

    existing_slugs: set[str] = set()
    workbook_by_workspace: dict[str, str] = {}

    for workspace in workspace_rows:
        slug = _unique_slug(str(workspace["name"]), existing_slugs)
        existing_slugs.add(slug)
        connection.execute(
            sa.text("UPDATE workspaces SET slug = :slug, owner_id = :owner_id WHERE id = :id"),
            {"slug": slug, "owner_id": SYSTEM_USER_ID, "id": workspace["id"]},
        )

    op.alter_column("workspaces", "slug", nullable=False)
    op.alter_column("workspaces", "owner_id", nullable=False)
    op.create_index("ix_workspaces_slug", "workspaces", ["slug"], unique=True)
    op.create_index("ix_workspaces_owner_id", "workspaces", ["owner_id"], unique=False)
    op.create_foreign_key(
        "fk_workspaces_owner_id_users",
        "workspaces",
        "users",
        ["owner_id"],
        ["id"],
        ondelete="RESTRICT",
    )

    op.create_table(
        "workspace_members",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("workspace_id", sa.String(length=32), nullable=False),
        sa.Column("user_id", sa.String(length=32), nullable=False),
        sa.Column("role", sa.String(length=24), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["user_id"], ["users.id"], ondelete="CASCADE"),
        sa.ForeignKeyConstraint(["workspace_id"], ["workspaces.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("workspace_id", "user_id", name="uq_workspace_members_workspace_user"),
    )
    op.create_index("ix_workspace_members_workspace_id", "workspace_members", ["workspace_id"], unique=False)
    op.create_index("ix_workspace_members_user_id", "workspace_members", ["user_id"], unique=False)

    op.create_table(
        "workbooks",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("workspace_id", sa.String(length=32), nullable=False),
        sa.Column("name", sa.String(length=120), nullable=False),
        sa.Column("description", sa.Text(), nullable=True),
        sa.Column("position", sa.Integer(), nullable=False),
        sa.Column("created_by", sa.String(length=32), nullable=False),
        sa.Column("updated_by", sa.String(length=32), nullable=False),
        sa.Column("archived", sa.Boolean(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["created_by"], ["users.id"], ondelete="RESTRICT"),
        sa.ForeignKeyConstraint(["updated_by"], ["users.id"], ondelete="RESTRICT"),
        sa.ForeignKeyConstraint(["workspace_id"], ["workspaces.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_workbooks_workspace_id", "workbooks", ["workspace_id"], unique=False)
    op.create_index("ix_workbooks_position", "workbooks", ["position"], unique=False)

    for workspace in workspace_rows:
        workbook_id = f"wb_{uuid4().hex[:8]}"
        workbook_by_workspace[str(workspace["id"])] = workbook_id
        connection.execute(
            sa.text(
                """
                INSERT INTO workspace_members (id, workspace_id, user_id, role, created_at)
                VALUES (:id, :workspace_id, :user_id, :role, :created_at)
                """
            ),
            {
                "id": f"wm_{uuid4().hex[:8]}",
                "workspace_id": workspace["id"],
                "user_id": SYSTEM_USER_ID,
                "role": "owner",
                "created_at": workspace["created_at"] or now,
            },
        )
        connection.execute(
            sa.text(
                """
                INSERT INTO workbooks (
                    id,
                    workspace_id,
                    name,
                    description,
                    position,
                    created_by,
                    updated_by,
                    archived,
                    created_at,
                    updated_at
                )
                VALUES (
                    :id,
                    :workspace_id,
                    :name,
                    :description,
                    :position,
                    :created_by,
                    :updated_by,
                    :archived,
                    :created_at,
                    :updated_at
                )
                """
            ),
            {
                "id": workbook_id,
                "workspace_id": workspace["id"],
                "name": f"{workspace['name']} workbook",
                "description": workspace["description"] if "description" in workspace else None,
                "position": 0,
                "created_by": SYSTEM_USER_ID,
                "updated_by": SYSTEM_USER_ID,
                "archived": False,
                "created_at": workspace["created_at"] or now,
                "updated_at": workspace["updated_at"] or now,
            },
        )

    op.create_table(
        "sheet_columns",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("sheet_id", sa.String(length=32), nullable=False),
        sa.Column("key", sa.String(length=120), nullable=False),
        sa.Column("title", sa.String(length=120), nullable=False),
        sa.Column("type", sa.String(length=32), nullable=False),
        sa.Column("position", sa.Integer(), nullable=False),
        sa.Column("width", sa.Integer(), nullable=True),
        sa.Column("nullable", sa.Boolean(), nullable=False),
        sa.Column("editable", sa.Boolean(), nullable=False),
        sa.Column("computed", sa.Boolean(), nullable=False),
        sa.Column("expression", sa.Text(), nullable=True),
        sa.Column("settings_json", sa.JSON(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["sheet_id"], ["sheets.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("sheet_id", "key", name="uq_sheet_columns_sheet_key"),
    )
    op.create_index("ix_sheet_columns_sheet_id", "sheet_columns", ["sheet_id"], unique=False)

    op.create_table(
        "sheet_records",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("sheet_id", sa.String(length=32), nullable=False),
        sa.Column("position", sa.Integer(), nullable=False),
        sa.Column("values_json", sa.JSON(), nullable=False),
        sa.Column("display_values_json", sa.JSON(), nullable=True),
        sa.Column("created_by", sa.String(length=32), nullable=False),
        sa.Column("updated_by", sa.String(length=32), nullable=False),
        sa.Column("archived", sa.Boolean(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["created_by"], ["users.id"], ondelete="RESTRICT"),
        sa.ForeignKeyConstraint(["sheet_id"], ["sheets.id"], ondelete="CASCADE"),
        sa.ForeignKeyConstraint(["updated_by"], ["users.id"], ondelete="RESTRICT"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_sheet_records_sheet_id", "sheet_records", ["sheet_id"], unique=False)

    op.create_table(
        "sheet_views",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("sheet_id", sa.String(length=32), nullable=False),
        sa.Column("key", sa.String(length=120), nullable=False),
        sa.Column("name", sa.String(length=120), nullable=False),
        sa.Column("type", sa.String(length=24), nullable=False),
        sa.Column("is_default", sa.Boolean(), nullable=False),
        sa.Column("created_by", sa.String(length=32), nullable=False),
        sa.Column("config_json", sa.JSON(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.Column("updated_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["created_by"], ["users.id"], ondelete="RESTRICT"),
        sa.ForeignKeyConstraint(["sheet_id"], ["sheets.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
        sa.UniqueConstraint("sheet_id", "key", name="uq_sheet_views_sheet_key"),
    )
    op.create_index("ix_sheet_views_sheet_id", "sheet_views", ["sheet_id"], unique=False)

    op.create_table(
        "activity_logs",
        sa.Column("id", sa.String(length=32), nullable=False),
        sa.Column("workspace_id", sa.String(length=32), nullable=True),
        sa.Column("workbook_id", sa.String(length=32), nullable=True),
        sa.Column("sheet_id", sa.String(length=32), nullable=True),
        sa.Column("record_id", sa.String(length=32), nullable=True),
        sa.Column("user_id", sa.String(length=32), nullable=True),
        sa.Column("action_type", sa.String(length=64), nullable=False),
        sa.Column("payload_json", sa.JSON(), nullable=False),
        sa.Column("created_at", sa.DateTime(timezone=True), nullable=False),
        sa.ForeignKeyConstraint(["record_id"], ["sheet_records.id"], ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["sheet_id"], ["sheets.id"], ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["user_id"], ["users.id"], ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["workbook_id"], ["workbooks.id"], ondelete="SET NULL"),
        sa.ForeignKeyConstraint(["workspace_id"], ["workspaces.id"], ondelete="CASCADE"),
        sa.PrimaryKeyConstraint("id"),
    )
    op.create_index("ix_activity_logs_workspace_id", "activity_logs", ["workspace_id"], unique=False)
    op.create_index("ix_activity_logs_workbook_id", "activity_logs", ["workbook_id"], unique=False)
    op.create_index("ix_activity_logs_sheet_id", "activity_logs", ["sheet_id"], unique=False)
    op.create_index("ix_activity_logs_record_id", "activity_logs", ["record_id"], unique=False)
    op.create_index("ix_activity_logs_user_id", "activity_logs", ["user_id"], unique=False)
    op.create_index("ix_activity_logs_action_type", "activity_logs", ["action_type"], unique=False)

    op.add_column("sheets", sa.Column("workbook_id", sa.String(length=32), nullable=True))
    op.add_column("sheets", sa.Column("key", sa.String(length=120), nullable=True))
    op.add_column(
        "sheets",
        sa.Column("kind", sa.String(length=24), nullable=False, server_default="data"),
    )
    op.add_column("sheets", sa.Column("source_sheet_id", sa.String(length=32), nullable=True))
    op.add_column(
        "sheets",
        sa.Column(
            "config_json",
            sa.JSON(),
            nullable=False,
            server_default=sa.text("'{}'::json"),
        ),
    )

    sheet_rows = connection.execute(
        sa.text(
            """
            SELECT id, workspace_id, name, icon, position, columns, rows, created_at, updated_at
            FROM sheets
            ORDER BY workspace_id, position, created_at, id
            """
        )
    ).mappings().all()
    keys_by_workbook: dict[str, set[str]] = {}

    for sheet in sheet_rows:
        workbook_id = workbook_by_workspace[str(sheet["workspace_id"])]
        workbook_keys = keys_by_workbook.setdefault(workbook_id, set())
        sheet_key = _unique_key(str(sheet["name"]), workbook_keys)
        workbook_keys.add(sheet_key)
        created_at = sheet["created_at"] or now
        updated_at = sheet["updated_at"] or now

        connection.execute(
            sa.text(
                """
                UPDATE sheets
                SET workbook_id = :workbook_id,
                    key = :key,
                    kind = :kind,
                    source_sheet_id = NULL,
                    config_json = CAST(:config_json AS JSON)
                WHERE id = :id
                """
            ),
            {
                "workbook_id": workbook_id,
                "key": sheet_key,
                "kind": _sheet_kind_from_icon(sheet["icon"]),
                "config_json": json.dumps({}),
                "id": sheet["id"],
            },
        )

        for index, column in enumerate(sheet["columns"] or []):
            key = str(column.get("key") or f"column_{index + 1}")
            connection.execute(
                sa.text(
                    """
                    INSERT INTO sheet_columns (
                        id,
                        sheet_id,
                        key,
                        title,
                        type,
                        position,
                        width,
                        nullable,
                        editable,
                        computed,
                        expression,
                        settings_json,
                        created_at,
                        updated_at
                    )
                    VALUES (
                        :id,
                        :sheet_id,
                        :key,
                        :title,
                        :type,
                        :position,
                        :width,
                        :nullable,
                        :editable,
                        :computed,
                        :expression,
                        CAST(:settings_json AS JSON),
                        :created_at,
                        :updated_at
                    )
                    """
                ),
                {
                    "id": str(column.get("id") or f"col_{uuid4().hex[:8]}"),
                    "sheet_id": sheet["id"],
                    "key": key,
                    "title": str(column.get("label") or key.replace("_", " ").title()),
                    "type": str(column.get("data_type") or "text"),
                    "position": index,
                    "width": column.get("width"),
                    "nullable": bool(column.get("nullable", True)),
                    "editable": bool(column.get("editable", True)),
                    "computed": bool(column.get("computed", False)),
                    "expression": column.get("expression"),
                    "settings_json": json.dumps(
                        {
                            key_name: value
                            for key_name, value in column.items()
                            if key_name
                            not in {"id", "key", "label", "data_type", "width", "nullable", "editable", "computed", "expression"}
                        }
                    ),
                    "created_at": created_at,
                    "updated_at": updated_at,
                },
            )

        for index, record in enumerate(sheet["rows"] or []):
            row_payload = dict(record or {})
            row_id = str(row_payload.pop("id", f"rec_{uuid4().hex[:8]}"))
            connection.execute(
                sa.text(
                    """
                    INSERT INTO sheet_records (
                        id,
                        sheet_id,
                        position,
                        values_json,
                        display_values_json,
                        created_by,
                        updated_by,
                        archived,
                        created_at,
                        updated_at
                    )
                    VALUES (
                        :id,
                        :sheet_id,
                        :position,
                        CAST(:values_json AS JSON),
                        CAST(:display_values_json AS JSON),
                        :created_by,
                        :updated_by,
                        :archived,
                        :created_at,
                        :updated_at
                    )
                    """
                ),
                {
                    "id": row_id,
                    "sheet_id": sheet["id"],
                    "position": index,
                    "values_json": json.dumps(row_payload),
                    "display_values_json": None,
                    "created_by": SYSTEM_USER_ID,
                    "updated_by": SYSTEM_USER_ID,
                    "archived": False,
                    "created_at": created_at,
                    "updated_at": updated_at,
                },
            )

        connection.execute(
            sa.text(
                """
                INSERT INTO sheet_views (
                    id,
                    sheet_id,
                    key,
                    name,
                    type,
                    is_default,
                    created_by,
                    config_json,
                    created_at,
                    updated_at
                )
                VALUES (
                    :id,
                    :sheet_id,
                    :key,
                    :name,
                    :type,
                    :is_default,
                    :created_by,
                    CAST(:config_json AS JSON),
                    :created_at,
                    :updated_at
                )
                """
            ),
            {
                "id": f"view_{uuid4().hex[:8]}",
                "sheet_id": sheet["id"],
                "key": "grid_default",
                "name": "Grid",
                "type": "grid",
                "is_default": True,
                "created_by": SYSTEM_USER_ID,
                "config_json": json.dumps(_default_view_config()),
                "created_at": created_at,
                "updated_at": updated_at,
            },
        )

    inspector = sa.inspect(connection)
    for foreign_key in inspector.get_foreign_keys("sheets"):
        if foreign_key["constrained_columns"] == ["workspace_id"]:
            op.drop_constraint(foreign_key["name"], "sheets", type_="foreignkey")
            break

    op.drop_index("ix_sheets_workspace_id", table_name="sheets")
    op.drop_column("sheets", "workspace_id")
    op.drop_column("sheets", "icon")
    op.drop_column("sheets", "columns")
    op.drop_column("sheets", "rows")
    op.alter_column("sheets", "workbook_id", nullable=False)
    op.alter_column("sheets", "key", nullable=False)
    op.alter_column("sheets", "kind", server_default=None)
    op.alter_column("sheets", "config_json", server_default=None)
    op.create_index("ix_sheets_workbook_id", "sheets", ["workbook_id"], unique=False)
    op.create_index("ix_sheets_source_sheet_id", "sheets", ["source_sheet_id"], unique=False)
    op.create_foreign_key(
        "fk_sheets_workbook_id_workbooks",
        "sheets",
        "workbooks",
        ["workbook_id"],
        ["id"],
        ondelete="CASCADE",
    )
    op.create_foreign_key(
        "fk_sheets_source_sheet_id_sheets",
        "sheets",
        "sheets",
        ["source_sheet_id"],
        ["id"],
        ondelete="SET NULL",
    )
    op.create_unique_constraint("uq_sheets_workbook_key", "sheets", ["workbook_id", "key"])


def downgrade() -> None:
    connection = op.get_bind()
    now = _timestamp()

    op.add_column("sheets", sa.Column("workspace_id", sa.String(length=32), nullable=True))
    op.add_column("sheets", sa.Column("icon", sa.String(length=24), nullable=False, server_default="grid"))
    op.add_column(
        "sheets",
        sa.Column(
            "columns",
            sa.JSON(),
            nullable=False,
            server_default=sa.text("'[]'::json"),
        ),
    )
    op.add_column(
        "sheets",
        sa.Column(
            "rows",
            sa.JSON(),
            nullable=False,
            server_default=sa.text("'[]'::json"),
        ),
    )

    sheet_rows = connection.execute(
        sa.text(
            """
            SELECT sheets.id, sheets.name, sheets.kind, workbooks.workspace_id
            FROM sheets
            JOIN workbooks ON workbooks.id = sheets.workbook_id
            ORDER BY sheets.position, sheets.created_at, sheets.id
            """
        )
    ).mappings().all()

    for sheet in sheet_rows:
        columns = connection.execute(
            sa.text(
                """
                SELECT key, title, type, width
                FROM sheet_columns
                WHERE sheet_id = :sheet_id
                ORDER BY position, created_at, id
                """
            ),
            {"sheet_id": sheet["id"]},
        ).mappings().all()
        records = connection.execute(
            sa.text(
                """
                SELECT id, values_json
                FROM sheet_records
                WHERE sheet_id = :sheet_id AND archived = false
                ORDER BY position, created_at, id
                """
            ),
            {"sheet_id": sheet["id"]},
        ).mappings().all()

        connection.execute(
            sa.text(
                """
                UPDATE sheets
                SET workspace_id = :workspace_id,
                    icon = :icon,
                    columns = CAST(:columns AS JSON),
                    rows = CAST(:rows AS JSON)
                WHERE id = :id
                """
            ),
            {
                "workspace_id": sheet["workspace_id"],
                "icon": "grid" if sheet["kind"] == "data" else sheet["kind"],
                "columns": json.dumps(
                    [
                        {
                            "key": column["key"],
                            "label": column["title"],
                            "data_type": column["type"],
                            "width": column["width"],
                        }
                        for column in columns
                    ]
                ),
                "rows": json.dumps(
                    [
                        {"id": record["id"], **(record["values_json"] or {})}
                        for record in records
                    ]
                ),
                "id": sheet["id"],
            },
        )

    op.alter_column("sheets", "workspace_id", nullable=False)
    op.alter_column("sheets", "icon", server_default=None)
    op.alter_column("sheets", "columns", server_default=None)
    op.alter_column("sheets", "rows", server_default=None)
    op.create_index("ix_sheets_workspace_id", "sheets", ["workspace_id"], unique=False)
    op.create_foreign_key(
        "sheets_workspace_id_fkey",
        "sheets",
        "workspaces",
        ["workspace_id"],
        ["id"],
        ondelete="CASCADE",
    )

    op.drop_constraint("uq_sheets_workbook_key", "sheets", type_="unique")
    op.drop_constraint("fk_sheets_source_sheet_id_sheets", "sheets", type_="foreignkey")
    op.drop_constraint("fk_sheets_workbook_id_workbooks", "sheets", type_="foreignkey")
    op.drop_index("ix_sheets_source_sheet_id", table_name="sheets")
    op.drop_index("ix_sheets_workbook_id", table_name="sheets")
    op.drop_column("sheets", "config_json")
    op.drop_column("sheets", "source_sheet_id")
    op.drop_column("sheets", "kind")
    op.drop_column("sheets", "key")
    op.drop_column("sheets", "workbook_id")

    op.drop_index("ix_activity_logs_action_type", table_name="activity_logs")
    op.drop_index("ix_activity_logs_user_id", table_name="activity_logs")
    op.drop_index("ix_activity_logs_record_id", table_name="activity_logs")
    op.drop_index("ix_activity_logs_sheet_id", table_name="activity_logs")
    op.drop_index("ix_activity_logs_workbook_id", table_name="activity_logs")
    op.drop_index("ix_activity_logs_workspace_id", table_name="activity_logs")
    op.drop_table("activity_logs")

    op.drop_index("ix_sheet_views_sheet_id", table_name="sheet_views")
    op.drop_table("sheet_views")

    op.drop_index("ix_sheet_records_sheet_id", table_name="sheet_records")
    op.drop_table("sheet_records")

    op.drop_index("ix_sheet_columns_sheet_id", table_name="sheet_columns")
    op.drop_table("sheet_columns")

    op.drop_index("ix_workbooks_position", table_name="workbooks")
    op.drop_index("ix_workbooks_workspace_id", table_name="workbooks")
    op.drop_table("workbooks")

    op.drop_index("ix_workspace_members_user_id", table_name="workspace_members")
    op.drop_index("ix_workspace_members_workspace_id", table_name="workspace_members")
    op.drop_table("workspace_members")

    op.drop_constraint("fk_workspaces_owner_id_users", "workspaces", type_="foreignkey")
    op.drop_index("ix_workspaces_owner_id", table_name="workspaces")
    op.drop_index("ix_workspaces_slug", table_name="workspaces")
    op.drop_column("workspaces", "owner_id")
    op.drop_column("workspaces", "slug")

    op.drop_index("ix_users_email", table_name="users")
    op.drop_table("users")
