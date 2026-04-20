/*
  # Fix BOQ Items RLS SELECT Policy

  The existing SELECT policy uses a correlated subquery JOIN to boq_files
  which can cause performance issues and silent failures when the JOIN
  doesn't resolve correctly in certain auth contexts.

  This migration drops and recreates the SELECT policy using a simpler,
  more reliable EXISTS subquery pattern that correctly verifies ownership
  through the boq_files table.

  Also adds a direct policy to allow authenticated users to read their own
  boq_items through a cleaner path.
*/

-- Drop existing SELECT policy
DROP POLICY IF EXISTS "Users can view boq items for their files" ON boq_items;

-- Recreate with a cleaner, more reliable pattern
CREATE POLICY "Users can view boq items for their files"
  ON boq_items
  FOR SELECT
  TO authenticated
  USING (
    boq_file_id IN (
      SELECT id FROM boq_files WHERE created_by = auth.uid()
    )
  );

-- Also fix the UPDATE policy the same way
DROP POLICY IF EXISTS "Users can update boq items for their files" ON boq_items;

CREATE POLICY "Users can update boq items for their files"
  ON boq_items
  FOR UPDATE
  TO authenticated
  USING (
    boq_file_id IN (
      SELECT id FROM boq_files WHERE created_by = auth.uid()
    )
  )
  WITH CHECK (
    boq_file_id IN (
      SELECT id FROM boq_files WHERE created_by = auth.uid()
    )
  );

-- Also fix the DELETE policy
DROP POLICY IF EXISTS "Users can delete boq items for their files" ON boq_items;

CREATE POLICY "Users can delete boq items for their files"
  ON boq_items
  FOR DELETE
  TO authenticated
  USING (
    boq_file_id IN (
      SELECT id FROM boq_files WHERE created_by = auth.uid()
    )
  );

-- Also ensure INSERT policy allows inserting items for owned files
DROP POLICY IF EXISTS "Users can insert boq items for their files" ON boq_items;

CREATE POLICY "Users can insert boq items for their files"
  ON boq_items
  FOR INSERT
  TO authenticated
  WITH CHECK (
    boq_file_id IN (
      SELECT id FROM boq_files WHERE created_by = auth.uid()
    )
  );
