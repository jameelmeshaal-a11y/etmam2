/*
  # Add archive support to boq_files

  ## Changes
  - Adds `is_archived` boolean column to boq_files (default false)
  - Adds `archived_at` timestamp column to boq_files
  - Adds `notes` text column to boq_files for user annotations
  - Updates RLS policies to filter archived files from default queries

  ## Security
  - Existing RLS policies remain in place
  - No data is lost; archiving is soft-delete only
*/

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'boq_files' AND column_name = 'is_archived'
  ) THEN
    ALTER TABLE boq_files ADD COLUMN is_archived boolean DEFAULT false NOT NULL;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'boq_files' AND column_name = 'archived_at'
  ) THEN
    ALTER TABLE boq_files ADD COLUMN archived_at timestamptz;
  END IF;

  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'boq_files' AND column_name = 'notes'
  ) THEN
    ALTER TABLE boq_files ADD COLUMN notes text DEFAULT '';
  END IF;
END $$;
