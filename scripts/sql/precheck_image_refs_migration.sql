-- Precheck current alembic revision and file.image_refs presence
SELECT version_num FROM alembic_version;

SELECT column_name, data_type
FROM information_schema.columns
WHERE table_name='file' AND column_name='image_refs';
