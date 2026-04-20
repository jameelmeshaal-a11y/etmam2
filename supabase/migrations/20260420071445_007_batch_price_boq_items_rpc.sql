/*
  # Add batch_price_boq_items RPC Function

  This function accepts an array of pricing updates and applies them all in a
  single database transaction, replacing the previous pattern of one UPDATE
  per item which caused timeouts and browser crashes for large BOQ files.

  Parameters:
    - p_updates: JSON array of objects with fields:
        id (uuid), unit_rate (numeric), total_price (numeric),
        status (text), linked_rate_id (uuid), confidence (integer)

  Returns: count of rows updated
*/

CREATE OR REPLACE FUNCTION batch_price_boq_items(p_updates jsonb)
RETURNS integer
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_count integer := 0;
  v_row jsonb;
BEGIN
  FOR v_row IN SELECT * FROM jsonb_array_elements(p_updates)
  LOOP
    UPDATE boq_items SET
      unit_rate      = (v_row->>'unit_rate')::numeric,
      total_price    = (v_row->>'total_price')::numeric,
      status         = v_row->>'status',
      linked_rate_id = (v_row->>'linked_rate_id')::uuid,
      confidence     = (v_row->>'confidence')::integer,
      updated_at     = now()
    WHERE id = (v_row->>'id')::uuid;
    v_count := v_count + 1;
  END LOOP;
  RETURN v_count;
END;
$$;

GRANT EXECUTE ON FUNCTION batch_price_boq_items(jsonb) TO authenticated;
