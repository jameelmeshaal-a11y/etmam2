/*
  # Fix total_items to exclude descriptive and zero-quantity items

  ## Problem
  total_items in boq_files was counting ALL rows including descriptive items
  and zero-quantity rows. This caused the priced percentage to show incorrectly
  because descriptive items inflated the denominator but were never counted
  as "priced".

  ## Fix
  total_items now counts only "priceable" items:
    - status != 'descriptive'
    - quantity IS NOT NULL AND quantity > 0

  This matches exactly what BOQTable and BOQSummaryPage use on the frontend.

  ## Changes
  1. price_boq_file_stats_only — corrected total_items filter (no updated_at)
  2. Backfill all existing boq_files with corrected values
*/

CREATE OR REPLACE FUNCTION price_boq_file_stats_only(p_boq_file_id uuid)
RETURNS void
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_total  integer;
  v_priced integer;
BEGIN
  SELECT
    COUNT(*) FILTER (
      WHERE status != 'descriptive'
        AND quantity IS NOT NULL
        AND quantity > 0
    ),
    COUNT(*) FILTER (
      WHERE status != 'descriptive'
        AND quantity IS NOT NULL
        AND quantity > 0
        AND unit_rate IS NOT NULL
        AND unit_rate > 0
    )
  INTO v_total, v_priced
  FROM boq_items
  WHERE boq_file_id = p_boq_file_id;

  UPDATE boq_files
  SET total_items  = v_total,
      priced_items = v_priced
  WHERE id = p_boq_file_id;
END;
$$;

-- Backfill all existing boq_files with corrected stats
DO $$
DECLARE
  r record;
BEGIN
  FOR r IN SELECT id FROM boq_files LOOP
    PERFORM price_boq_file_stats_only(r.id);
  END LOOP;
END $$;
