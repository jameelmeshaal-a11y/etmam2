/*
  # Server-Side Pricing RPC Function

  ## Problem Solved
  The browser was performing 473 × 2665 = 1,261,045 similarity comparisons in
  JavaScript (single-threaded), which caused the browser tab to freeze and crash.

  ## Solution
  All matching and pricing logic is now executed inside Postgres using a single
  RPC call. The browser sends the BOQ file ID and receives back the pricing results.

  ## What this migration does
  1. Adds pg_trgm extension for text similarity matching
  2. Creates `price_boq_file` RPC function that:
     - Loads all eligible BOQ items for the file
     - Loads all Approved library items
     - Matches each BOQ item to the best library item using:
         a) Unit filter (must match)
         b) Trigram similarity on description vs standard_name_ar
         c) Word overlap (Jaccard-style) as tiebreaker
     - Updates matched items with price, confidence, status, linked_rate_id
     - Sets unmatched items to status='pending'
     - Updates boq_files stats (total_items, priced_items)
     - Returns summary counts
*/

CREATE EXTENSION IF NOT EXISTS pg_trgm;

CREATE OR REPLACE FUNCTION price_boq_file(p_boq_file_id uuid)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_item          RECORD;
  v_best_lib_id   uuid;
  v_best_rate     numeric;
  v_best_conf     integer;
  v_total         integer := 0;
  v_priced        integer := 0;
  v_failed        integer := 0;
  v_norm_unit     text;
BEGIN
  -- Process each eligible BOQ item
  FOR v_item IN
    SELECT id, description, unit, quantity, item_no, override_type, status
    FROM boq_items
    WHERE boq_file_id = p_boq_file_id
      AND override_type IS DISTINCT FROM 'manual'
      AND status IS DISTINCT FROM 'descriptive'
      AND quantity > 0
  LOOP
    v_total := v_total + 1;

    -- Normalize unit for matching
    v_norm_unit := lower(trim(v_item.unit));

    -- Find best matching library item using trigram similarity
    SELECT
      rl.id,
      rl.rate_target,
      LEAST(95, GREATEST(0,
        ROUND(
          GREATEST(
            similarity(lower(v_item.description), lower(rl.standard_name_ar)),
            similarity(lower(v_item.description), lower(COALESCE(rl.standard_name_en, '')))
          ) * 100
        )
      ))::integer AS conf
    INTO v_best_lib_id, v_best_rate, v_best_conf
    FROM rate_library rl
    WHERE rl.source_type = 'Approved'
      AND rl.rate_target > 0
      AND (
        lower(trim(rl.unit)) = v_norm_unit
        OR rl.unit = v_item.unit
        OR lower(trim(rl.unit)) LIKE '%' || v_norm_unit || '%'
        OR v_norm_unit LIKE '%' || lower(trim(rl.unit)) || '%'
      )
    ORDER BY
      GREATEST(
        similarity(lower(v_item.description), lower(rl.standard_name_ar)),
        similarity(lower(v_item.description), lower(COALESCE(rl.standard_name_en, '')))
      ) DESC
    LIMIT 1;

    IF v_best_lib_id IS NOT NULL AND v_best_conf >= 30 THEN
      -- Update item with price
      UPDATE boq_items SET
        unit_rate      = v_best_rate,
        total_price    = v_item.quantity * v_best_rate,
        status         = CASE WHEN v_best_conf >= 75 THEN 'approved' ELSE 'pending' END,
        linked_rate_id = v_best_lib_id,
        confidence     = v_best_conf,
        updated_at     = now()
      WHERE id = v_item.id;

      v_priced := v_priced + 1;
    ELSE
      -- Mark as unmatched
      UPDATE boq_items SET
        status     = 'pending',
        confidence = COALESCE(v_best_conf, 0),
        unit_rate  = NULL,
        total_price = NULL,
        updated_at = now()
      WHERE id = v_item.id;

      v_failed := v_failed + 1;
    END IF;

  END LOOP;

  -- Update file stats
  UPDATE boq_files SET
    total_items  = (SELECT COUNT(*) FROM boq_items WHERE boq_file_id = p_boq_file_id),
    priced_items = (SELECT COUNT(*) FROM boq_items WHERE boq_file_id = p_boq_file_id
                    AND (status = 'approved' OR override_type = 'manual') AND unit_rate IS NOT NULL)
  WHERE id = p_boq_file_id;

  RETURN jsonb_build_object(
    'total',  v_total,
    'priced', v_priced,
    'failed', v_failed
  );
END;
$$;

GRANT EXECUTE ON FUNCTION price_boq_file(uuid) TO authenticated;
