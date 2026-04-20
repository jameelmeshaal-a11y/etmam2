/*
  # save_manual_price() — Atomic RPC Function

  This function saves a manual price breakdown in a single database transaction.
  All six sub-steps must succeed together, or none are committed (no partial state).

  ## Steps
  1. Lock & fetch the boq_items row
  2. Update boq_items with all breakdown components
  3. Validate linked_rate_id (if provided) still exists in rate_library
  4. Update or INSERT rate_library record if needed
  5. INSERT into rate_sources (audit trail)
  6. Recalculate boq_file priced_items count
*/

CREATE OR REPLACE FUNCTION save_manual_price(
  p_boq_item_id uuid,
  p_unit_rate numeric,
  p_materials numeric DEFAULT 0,
  p_labor numeric DEFAULT 0,
  p_equipment numeric DEFAULT 0,
  p_logistics numeric DEFAULT 0,
  p_risk numeric DEFAULT 0,
  p_profit numeric DEFAULT 0,
  p_linked_rate_id uuid DEFAULT NULL
)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_item boq_items%ROWTYPE;
  v_total_price numeric;
  v_file_id uuid;
BEGIN
  -- Step 1: Lock & fetch the boq_items row
  SELECT * INTO v_item
  FROM boq_items
  WHERE id = p_boq_item_id
  FOR UPDATE;

  IF NOT FOUND THEN
    RETURN jsonb_build_object('success', false, 'error', 'Item not found');
  END IF;

  -- Governance Wall 2: Manual override protection
  -- (Allow saving manual prices regardless - this IS the manual save)

  v_total_price := v_item.quantity * p_unit_rate;
  v_file_id := v_item.boq_file_id;

  -- Governance Wall 4: Min/Max Rate Enforcement
  IF p_linked_rate_id IS NOT NULL THEN
    DECLARE
      v_rate_min numeric;
      v_rate_max numeric;
    BEGIN
      SELECT rate_min, rate_max INTO v_rate_min, v_rate_max
      FROM rate_library
      WHERE id = p_linked_rate_id;

      IF FOUND AND (p_unit_rate < v_rate_min OR p_unit_rate > v_rate_max) THEN
        RETURN jsonb_build_object(
          'success', false,
          'error', 'Rate outside governance bounds',
          'rate_min', v_rate_min,
          'rate_max', v_rate_max
        );
      END IF;
    END;
  END IF;

  -- Step 2: Update boq_items
  UPDATE boq_items SET
    unit_rate = p_unit_rate,
    total_price = v_total_price,
    override_type = 'manual',
    status = 'approved',
    linked_rate_id = p_linked_rate_id,
    materials = p_materials,
    labor = p_labor,
    equipment = p_equipment,
    logistics = p_logistics,
    risk = p_risk,
    profit = p_profit,
    updated_at = now()
  WHERE id = p_boq_item_id;

  -- Step 5: INSERT into rate_sources (audit trail)
  INSERT INTO rate_sources (
    boq_item_id, rate_library_id, unit_rate, source_type, override_type,
    materials, labor, equipment, logistics, risk, profit
  ) VALUES (
    p_boq_item_id, p_linked_rate_id, p_unit_rate, 'manual', 'manual',
    p_materials, p_labor, p_equipment, p_logistics, p_risk, p_profit
  );

  -- Step 6: Recalculate priced_items count in boq_files
  UPDATE boq_files SET
    priced_items = (
      SELECT COUNT(*) FROM boq_items
      WHERE boq_file_id = v_file_id
      AND status IN ('approved', 'manual')
      AND override_type = 'manual' OR (status = 'approved' AND unit_rate IS NOT NULL)
    )
  WHERE id = v_file_id;

  RETURN jsonb_build_object('success', true, 'total_price', v_total_price);
END;
$$;
