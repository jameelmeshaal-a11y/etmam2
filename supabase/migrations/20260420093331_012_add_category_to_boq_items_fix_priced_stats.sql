/*
  # Add category column to boq_items & fix priced_items stats logic

  ## Changes

  1. New Column
     - `boq_items.category` (text, nullable) — copied from linked rate library item when priced

  2. Fix priced_items calculation
     - Old logic: only counted approved + manual items
     - New logic: any priceable item (not descriptive, qty > 0) with unit_rate > 0
       This matches what BOQTable shows as "نسبة التسعير"

  3. Updated price_boq_file_stats_only function to use corrected logic
*/

DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM information_schema.columns
    WHERE table_name = 'boq_items' AND column_name = 'category'
  ) THEN
    ALTER TABLE boq_items ADD COLUMN category text;
  END IF;
END $$;

CREATE OR REPLACE FUNCTION price_boq_file_stats_only(p_boq_file_id uuid)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_total integer;
  v_priced integer;
BEGIN
  SELECT
    COUNT(*),
    COUNT(*) FILTER (
      WHERE status != 'descriptive'
        AND (quantity IS NULL OR quantity > 0)
        AND unit_rate IS NOT NULL
        AND unit_rate > 0
    )
  INTO v_total, v_priced
  FROM boq_items
  WHERE boq_file_id = p_boq_file_id;

  UPDATE boq_files
  SET total_items   = v_total,
      priced_items  = v_priced,
      updated_at    = now()
  WHERE id = p_boq_file_id;
END;
$$;

GRANT EXECUTE ON FUNCTION price_boq_file_stats_only(uuid) TO authenticated;
