/*
  # Add Foreign Key: boq_items.linked_rate_id → rate_library.id

  ## Problem
  The boq_items table has a linked_rate_id column but no foreign key constraint
  to the rate_library table. Without this FK, PostgREST (Supabase's API layer)
  cannot resolve the joined query `matched_library_item:rate_library(*)` used
  in BOQTable.tsx, causing the entire SELECT query to silently fail and return
  no rows — making all BOQ items invisible after upload.

  ## Changes
  - Adds FK constraint: boq_items.linked_rate_id → rate_library(id) ON DELETE SET NULL
  - This enables PostgREST to resolve the relationship for joined queries

  ## Notes
  - ON DELETE SET NULL ensures that if a rate library item is deleted, the
    boq_item's linked_rate_id is set to NULL rather than causing a cascade delete
  - Existing rows with NULL linked_rate_id are unaffected
*/

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.table_constraints
    WHERE constraint_name = 'boq_items_linked_rate_id_fkey'
    AND table_name = 'boq_items'
  ) THEN
    ALTER TABLE boq_items
      ADD CONSTRAINT boq_items_linked_rate_id_fkey
      FOREIGN KEY (linked_rate_id)
      REFERENCES rate_library(id)
      ON DELETE SET NULL;
  END IF;
END $$;
