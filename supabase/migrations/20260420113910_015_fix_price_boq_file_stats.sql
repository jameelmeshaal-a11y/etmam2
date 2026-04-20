/*
  # Fix price_boq_file RPC to use corrected stats logic

  The price_boq_file function was updating total_items with COUNT(*) (all rows)
  and priced_items only counting approved/manual. Now both use the same
  "priceable" definition: status != 'descriptive' AND quantity > 0.
*/

CREATE OR REPLACE FUNCTION price_boq_file(p_boq_file_id uuid, p_city text DEFAULT 'riyadh')
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_item         record;
  v_rate_row     record;
  v_total        integer := 0;
  v_priced       integer := 0;
  v_failed       integer := 0;
  v_unit_rate    numeric;
  v_total_price  numeric;
BEGIN
  FOR v_item IN
    SELECT id, description, quantity, status, unit, category
    FROM boq_items
    WHERE boq_file_id = p_boq_file_id
      AND status NOT IN ('descriptive', 'approved', 'manual')
      AND quantity IS NOT NULL
      AND quantity > 0
  LOOP
    v_total := v_total + 1;

    SELECT rate_target
    INTO v_unit_rate
    FROM rate_library
    WHERE is_locked = false
    ORDER BY rate_target ASC
    LIMIT 1;

    IF NOT FOUND OR v_unit_rate IS NULL THEN
      v_failed := v_failed + 1;
      CONTINUE;
    END IF;

    v_total_price := v_unit_rate * v_item.quantity;

    UPDATE boq_items
    SET unit_rate   = v_unit_rate,
        total_price = v_total_price,
        status      = 'approved'
    WHERE id = v_item.id;

    v_priced := v_priced + 1;
  END LOOP;

  -- Update file stats using priceable definition (excludes descriptive + zero qty)
  UPDATE boq_files
  SET total_items  = (
        SELECT COUNT(*)
        FROM boq_items
        WHERE boq_file_id = p_boq_file_id
          AND status != 'descriptive'
          AND quantity IS NOT NULL
          AND quantity > 0
      ),
      priced_items = (
        SELECT COUNT(*)
        FROM boq_items
        WHERE boq_file_id = p_boq_file_id
          AND status != 'descriptive'
          AND quantity IS NOT NULL
          AND quantity > 0
          AND unit_rate IS NOT NULL
          AND unit_rate > 0
      )
  WHERE id = p_boq_file_id;

  RETURN jsonb_build_object(
    'total',  v_total,
    'priced', v_priced,
    'failed', v_failed
  );
END;
$$;
