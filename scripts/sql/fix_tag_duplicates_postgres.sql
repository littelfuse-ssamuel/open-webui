-- PostgreSQL tag cleanup for Open WebUI
-- Purpose:
-- 1) Remove invalid duplicate rows by (id, user_id)
-- 2) Enforce current tag key shape: PRIMARY KEY (id, user_id)
-- 3) Keep duplicate id values across different users (this is valid)

BEGIN;

-- One-time safety backup of current data
CREATE TABLE IF NOT EXISTS tag_backup_before_dedupe AS
SELECT * FROM tag;

-- Abort if key columns are NULL; fix manually first if this triggers
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM tag WHERE id IS NULL OR user_id IS NULL) THEN
        RAISE EXCEPTION 'tag contains NULL in id or user_id; aborting cleanup';
    END IF;
END
$$;

-- Keep one row per (id, user_id), delete extras
WITH ranked AS (
    SELECT
        ctid,
        ROW_NUMBER() OVER (
            PARTITION BY id, user_id
            ORDER BY ctid
        ) AS rn
    FROM tag
)
DELETE FROM tag t
USING ranked r
WHERE t.ctid = r.ctid
  AND r.rn > 1;

-- Drop legacy id-only constraints/indexes if present
ALTER TABLE tag DROP CONSTRAINT IF EXISTS tag_pkey;
ALTER TABLE tag DROP CONSTRAINT IF EXISTS pk_id;
ALTER TABLE tag DROP CONSTRAINT IF EXISTS uq_tag_id;
ALTER TABLE tag DROP CONSTRAINT IF EXISTS tag_id_key;
DROP INDEX IF EXISTS tag_id;
DROP INDEX IF EXISTS tag_id_key;

-- Enforce current schema key
DO $$
BEGIN
    IF NOT EXISTS (
        SELECT 1
        FROM pg_constraint
        WHERE conrelid = 'tag'::regclass
          AND contype = 'p'
          AND conname = 'pk_id_user_id'
    ) THEN
        ALTER TABLE tag
            ADD CONSTRAINT pk_id_user_id PRIMARY KEY (id, user_id);
    END IF;
END
$$;

COMMIT;

-- Validation: should return 0 rows
SELECT id, user_id, COUNT(*) AS row_count
FROM tag
GROUP BY id, user_id
HAVING COUNT(*) > 1;

-- Informational only: duplicate id across different users is expected/allowed.
SELECT id, COUNT(*) AS total_rows, COUNT(DISTINCT user_id) AS distinct_users
FROM tag
GROUP BY id
HAVING COUNT(*) > 1
ORDER BY total_rows DESC, id;

