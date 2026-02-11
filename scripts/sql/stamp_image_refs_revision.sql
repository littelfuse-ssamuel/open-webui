-- Safely stamp alembic revision from c440947495f3 to 8257b99d21e3 only if image_refs exists
BEGIN;

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1
    FROM information_schema.columns
    WHERE table_name='file' AND column_name='image_refs'
  ) THEN
    RAISE EXCEPTION 'file.image_refs does not exist; refusing to stamp';
  END IF;
END
$$;

UPDATE alembic_version
SET version_num = '8257b99d21e3'
WHERE version_num = 'c440947495f3';

COMMIT;
