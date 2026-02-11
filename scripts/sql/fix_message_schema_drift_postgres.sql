-- Repair Open WebUI message table schema drift on PostgreSQL
-- Safe to run multiple times.

BEGIN;

ALTER TABLE message ADD COLUMN IF NOT EXISTS reply_to_id TEXT;
ALTER TABLE message ADD COLUMN IF NOT EXISTS parent_id TEXT;
ALTER TABLE message ADD COLUMN IF NOT EXISTS is_pinned BOOLEAN;
ALTER TABLE message ADD COLUMN IF NOT EXISTS pinned_at BIGINT;
ALTER TABLE message ADD COLUMN IF NOT EXISTS pinned_by TEXT;

-- Normalize is_pinned semantics expected by app
UPDATE message SET is_pinned = FALSE WHERE is_pinned IS NULL;
ALTER TABLE message ALTER COLUMN is_pinned SET DEFAULT FALSE;
ALTER TABLE message ALTER COLUMN is_pinned SET NOT NULL;

COMMIT;

-- Verification
SELECT column_name, data_type, is_nullable, column_default
FROM information_schema.columns
WHERE table_name = 'message'
  AND column_name IN ('reply_to_id','parent_id','is_pinned','pinned_at','pinned_by')
ORDER BY column_name;
