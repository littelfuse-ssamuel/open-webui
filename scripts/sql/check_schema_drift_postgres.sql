-- Open WebUI PostgreSQL schema drift audit (read-only)
-- Lists missing expected columns in key tables used by Channels, Files, and Documents/Knowledge.

WITH expected(table_name, column_name) AS (
  VALUES
    -- message/channels
    ('message','id'),
    ('message','user_id'),
    ('message','channel_id'),
    ('message','reply_to_id'),
    ('message','parent_id'),
    ('message','is_pinned'),
    ('message','pinned_at'),
    ('message','pinned_by'),
    ('message','content'),
    ('message','data'),
    ('message','meta'),
    ('message','created_at'),
    ('message','updated_at'),

    ('channel','id'),
    ('channel','user_id'),
    ('channel','name'),
    ('channel','description'),
    ('channel','type'),
    ('channel','data'),
    ('channel','meta'),
    ('channel','access_control'),
    ('channel','created_at'),
    ('channel','updated_at'),

    ('channel_member','id'),
    ('channel_member','channel_id'),
    ('channel_member','user_id'),
    ('channel_member','status'),
    ('channel_member','is_active'),
    ('channel_member','is_channel_muted'),
    ('channel_member','is_channel_pinned'),
    ('channel_member','data'),
    ('channel_member','meta'),
    ('channel_member','joined_at'),
    ('channel_member','left_at'),
    ('channel_member','last_read_at'),
    ('channel_member','created_at'),
    ('channel_member','updated_at'),

    -- files/chat-files
    ('file','id'),
    ('file','user_id'),
    ('file','hash'),
    ('file','filename'),
    ('file','path'),
    ('file','data'),
    ('file','meta'),
    ('file','access_control'),
    ('file','created_at'),
    ('file','updated_at'),
    ('file','image_refs'),

    ('chat_file','id'),
    ('chat_file','user_id'),
    ('chat_file','chat_id'),
    ('chat_file','file_id'),
    ('chat_file','message_id'),
    ('chat_file','created_at'),
    ('chat_file','updated_at'),

    -- documents / knowledge
    ('knowledge','id'),
    ('knowledge','user_id'),
    ('knowledge','name'),
    ('knowledge','description'),
    ('knowledge','data'),
    ('knowledge','meta'),
    ('knowledge','access_control'),
    ('knowledge','created_at'),
    ('knowledge','updated_at'),

    ('knowledge_file','id'),
    ('knowledge_file','knowledge_id'),
    ('knowledge_file','file_id'),
    ('knowledge_file','user_id'),
    ('knowledge_file','created_at'),
    ('knowledge_file','updated_at'),

    -- settings/config surface often hit from admin
    ('config','id'),
    ('config','data'),
    ('config','version'),
    ('config','created_at'),
    ('config','updated_at')
)
SELECT e.table_name, e.column_name
FROM expected e
LEFT JOIN information_schema.columns c
  ON c.table_name = e.table_name
 AND c.column_name = e.column_name
WHERE c.column_name IS NULL
ORDER BY e.table_name, e.column_name;

-- Current alembic revision
SELECT version_num FROM alembic_version;
