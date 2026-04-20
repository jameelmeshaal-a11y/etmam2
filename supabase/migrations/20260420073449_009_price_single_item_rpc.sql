/*
  # Price Single BOQ Item RPC

  ## Purpose
  Prices a single BOQ item server-side using pg_trgm similarity.
  This allows the frontend to price items one-by-one with real progress updates,
  avoiding the statement timeout that occurs when pricing hundreds of items in one call.

  ## What this creates
  - `price_single_boq_item(p_item_id uuid)` - Prices one item and returns the result

  ## Security
  - SECURITY DEFINER runs as the function owner
  - Only authenticated users can execute
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

  IF v_item.quantity IS NULL OR v_item.quantity <= 0 THEN
    RETURN jsonb_build_object('success', false, 'reason', 'no_quantity', 'skipped', true);
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
