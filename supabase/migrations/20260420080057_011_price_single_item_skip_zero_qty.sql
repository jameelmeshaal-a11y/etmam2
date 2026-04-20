/*
  # Update price_single_boq_item to skip zero-qty items

  ## Purpose
  Items with quantity = 0 or NULL should be treated as descriptive headers
  and skipped during pricing. This prevents them from showing as "unpriced".

  ## Changes
  - Updates the price_single_boq_item RPC to mark qty=0 items as descriptive
    and return skipped=true
*/

CREATE OR REPLACE FUNCTION price_single_boq_item(p_item_id uuid)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_item          RECORD;
  v_best_lib_id   uuid;
  v_best_rate     numeric;
  v_best_conf     integer;
  v_norm_unit     text;
BEGIN
  SELECT id, description, unit, quantity, item_no, override_type, status, boq_file_id
  INTO v_item
  FROM boq_items
  WHERE id = p_item_id;

  IF NOT FOUND THEN
    RETURN jsonb_build_object('success', false, 'reason', 'not_found');
  END IF;

  IF v_item.override_type = 'manual' THEN
    RETURN jsonb_build_object('success', false, 'reason', 'manual_override', 'skipped', true);
  END IF;

  IF v_item.status = 'descriptive' THEN
    RETURN jsonb_build_object('success', false, 'reason', 'descriptive', 'skipped', true);
  END IF;

  -- Items with zero or null quantity are structural/header rows → mark as descriptive
  IF v_item.quantity IS NULL OR v_item.quantity <= 0 THEN
    UPDATE boq_items SET status = 'descriptive', updated_at = now() WHERE id = p_item_id;
    RETURN jsonb_build_object('success', false, 'reason', 'zero_qty_descriptive', 'skipped', true);
  END IF;

  v_norm_unit := lower(trim(v_item.unit));

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
    UPDATE boq_items SET
      unit_rate      = v_best_rate,
      total_price    = v_item.quantity * v_best_rate,
      status         = CASE WHEN v_best_conf >= 75 THEN 'approved' ELSE 'pending' END,
      linked_rate_id = v_best_lib_id,
      confidence     = v_best_conf,
      updated_at     = now()
    WHERE id = p_item_id;

    RETURN jsonb_build_object(
      'success',     true,
      'unit_rate',   v_best_rate,
      'total_price', v_item.quantity * v_best_rate,
      'confidence',  v_best_conf,
      'lib_id',      v_best_lib_id,
      'status',      CASE WHEN v_best_conf >= 75 THEN 'approved' ELSE 'pending' END
    );
  ELSE
    UPDATE boq_items SET
      status      = 'pending',
      confidence  = COALESCE(v_best_conf, 0),
      unit_rate   = NULL,
      total_price = NULL,
      updated_at  = now()
    WHERE id = p_item_id;

    RETURN jsonb_build_object(
      'success',    false,
      'reason',     'no_match',
      'confidence', COALESCE(v_best_conf, 0)
    );
  END IF;
END;
$$;

GRANT EXECUTE ON FUNCTION price_single_boq_item(uuid) TO authenticated;
