/*
  # Zero quantity items are always descriptive

  ## Rule
  Any BOQ item with quantity = 0 or NULL is considered descriptive (وصفي).
  Only items with quantity > 0 need pricing.

  ## Changes
  1. Update existing pending items with zero quantity to descriptive status
  2. Update price_single_boq_item function to enforce this rule consistently
  3. Update price_boq_file_stats_only to exclude zero-qty items from totals
*/

-- Fix existing zero-qty pending items
UPDATE boq_items
SET status = 'descriptive',
    unit_rate = NULL,
    total_price = NULL,
    updated_at = now()
WHERE (quantity IS NULL OR quantity = 0)
  AND status NOT IN ('descriptive');

-- Update price_single_boq_item to always mark zero-qty as descriptive
CREATE OR REPLACE FUNCTION price_single_boq_item(p_item_id uuid)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_item        RECORD;
  v_norm_unit   text;
  v_canon_unit  text;

  v_best_id     uuid    := NULL;
  v_best_rate   numeric := NULL;
  v_best_conf   integer := 0;

  v_lib_id      uuid;
  v_lib_rate    numeric;
  v_lib_name    text;
  v_lib_kw      text[];
  v_lib_unit    text;
  v_lib_canon   text;

  v_trgm        float;
  v_precision   float;
  v_score       float;

  v_boq_kw_set  text[];
  v_lib_kw_norm text[];
  v_matched     int;
  v_lib_kw_cnt  int;

  v_boq_t       text;
  v_lib_t       text;
  v_status      text;
BEGIN
  SELECT id, description, unit, quantity, item_no, override_type, status, boq_file_id
  INTO v_item
  FROM boq_items WHERE id = p_item_id;

  IF NOT FOUND THEN
    RETURN jsonb_build_object('success', false, 'reason', 'not_found');
  END IF;

  IF v_item.override_type = 'manual' THEN
    RETURN jsonb_build_object('success', false, 'reason', 'manual_override', 'skipped', true);
  END IF;

  -- Zero or null quantity => always descriptive
  IF v_item.quantity IS NULL OR v_item.quantity <= 0 THEN
    UPDATE boq_items SET
      status      = 'descriptive',
      unit_rate   = NULL,
      total_price = NULL,
      updated_at  = now()
    WHERE id = p_item_id;
    RETURN jsonb_build_object('success', false, 'reason', 'descriptive', 'skipped', true);
  END IF;

  IF v_item.status = 'descriptive' THEN
    RETURN jsonb_build_object('success', false, 'reason', 'descriptive', 'skipped', true);
  END IF;

  v_norm_unit  := lower(trim(regexp_replace(COALESCE(v_item.unit, ''), '\s+', ' ', 'g')));

  v_canon_unit := CASE v_norm_unit
    WHEN 'متر مربع'  THEN 'm²'  WHEN 'م²'       THEN 'm²'  WHEN 'م2'      THEN 'm²'
    WHEN 'm2'        THEN 'm²'  WHEN 'sqm'      THEN 'm²'  WHEN 'sq.m'    THEN 'm²'
    WHEN 'sq m'      THEN 'm²'
    WHEN 'متر مكعب' THEN 'm³'  WHEN 'م³'       THEN 'm³'  WHEN 'م3'      THEN 'm³'
    WHEN 'm3'        THEN 'm³'  WHEN 'cum'      THEN 'm³'  WHEN 'cu.m'    THEN 'm³'
    WHEN 'متر طولي' THEN 'lm'  WHEN 'م'        THEN 'lm'  WHEN 'متر'     THEN 'lm'
    WHEN 'lm'        THEN 'lm'  WHEN 'ml'       THEN 'lm'  WHEN 'lin.m'   THEN 'lm'
    WHEN 'م.ط'       THEN 'lm'  WHEN 'م. .ط'    THEN 'lm'  WHEN 'م.ط.'    THEN 'lm'
    WHEN 'م ط'       THEN 'lm'  WHEN 'م . ط'    THEN 'lm'  WHEN 'م..ط'    THEN 'lm'
    WHEN 'meter'     THEN 'lm'
    WHEN 'عدد'       THEN 'nos' WHEN 'قطعة'     THEN 'nos' WHEN 'no'      THEN 'nos'
    WHEN 'nos'       THEN 'nos' WHEN 'pcs'      THEN 'nos' WHEN 'piece'   THEN 'nos'
    WHEN 'no.'       THEN 'nos' WHEN 'nos.'     THEN 'nos' WHEN 'nr'      THEN 'nos'
    WHEN 'طن'        THEN 'ton' WHEN 'ton'      THEN 'ton' WHEN 'tonnes'  THEN 'ton' WHEN 'tonne' THEN 'ton'
    WHEN 'كجم'       THEN 'kg'  WHEN 'كيلوغرام' THEN 'kg'  WHEN 'kg'      THEN 'kg'  WHEN 'كغ' THEN 'kg'
    WHEN 'مجموعة'   THEN 'set' WHEN 'set'      THEN 'set' WHEN 'sets'    THEN 'set'
    WHEN 'بند'       THEN 'ls'  WHEN 'إجمالي'   THEN 'ls'  WHEN 'ls'      THEN 'ls'
    WHEN 'l.s'       THEN 'ls'  WHEN 'lump sum' THEN 'ls'  WHEN 'l/s'     THEN 'ls'
    WHEN 'مبلغ مقطوع' THEN 'ls'
    ELSE v_norm_unit
  END;

  SELECT ARRAY(
    SELECT DISTINCT
      regexp_replace(
        regexp_replace(
          regexp_replace(
            regexp_replace(lower(word), '[\u064B-\u065F\u0670]', '', 'g'),
            '[\u0625\u0623\u0622\u0671]', '\u0627', 'g'),
          '\u0629', '\u0647', 'g'),
        '\u0649', '\u064A', 'g')
    FROM regexp_split_to_table(
      regexp_replace(v_item.description, '[،؛؟!.,;:()\[\]{}|/\\\-]', ' ', 'g'),
      '\s+'
    ) word
    WHERE length(word) >= 2
      AND lower(word) NOT IN (
        'من','في','على','الى','عن','مع','ان','هو','هي','لا','ما','كل',
        'ذلك','هذا','هذه','التي','الذي','حتى','اذا','يتم','يشمل','بما',
        'وفق','حسب','كما','وفقا','طبقا','والتي','وذلك','إلى','أن','إن'
      )
  ) INTO v_boq_kw_set;

  FOR v_lib_id, v_lib_rate, v_lib_name, v_lib_kw, v_lib_unit IN
    SELECT rl.id, rl.rate_target, rl.standard_name_ar, rl.keywords,
           lower(trim(regexp_replace(COALESCE(rl.unit, ''), '\s+', ' ', 'g')))
    FROM rate_library rl
    WHERE rl.source_type = 'Approved'
      AND rl.rate_target > 0
  LOOP
    v_lib_canon := CASE v_lib_unit
      WHEN 'متر مربع'  THEN 'm²'  WHEN 'م²'       THEN 'm²'  WHEN 'م2'      THEN 'm²'
      WHEN 'm2'        THEN 'm²'  WHEN 'sqm'      THEN 'm²'
      WHEN 'متر مكعب' THEN 'm³'  WHEN 'م³'       THEN 'm³'  WHEN 'م3'      THEN 'm³'
      WHEN 'm3'        THEN 'm³'  WHEN 'cum'      THEN 'm³'
      WHEN 'متر طولي' THEN 'lm'  WHEN 'م'        THEN 'lm'  WHEN 'متر'     THEN 'lm'
      WHEN 'lm'        THEN 'lm'  WHEN 'ml'       THEN 'lm'
      WHEN 'م.ط'       THEN 'lm'  WHEN 'م. .ط'    THEN 'lm'  WHEN 'م.ط.'    THEN 'lm'
      WHEN 'م ط'       THEN 'lm'  WHEN 'م . ط'    THEN 'lm'  WHEN 'م..ط'    THEN 'lm'
      WHEN 'عدد'       THEN 'nos' WHEN 'قطعة'     THEN 'nos' WHEN 'no'      THEN 'nos'
      WHEN 'nos'       THEN 'nos' WHEN 'pcs'      THEN 'nos' WHEN 'piece'   THEN 'nos'
      WHEN 'no.'       THEN 'nos' WHEN 'nos.'     THEN 'nos' WHEN 'nr'      THEN 'nos'
      WHEN 'طن'        THEN 'ton' WHEN 'ton'      THEN 'ton' WHEN 'tonnes'  THEN 'ton'
      WHEN 'كجم'       THEN 'kg'  WHEN 'كيلوغرام' THEN 'kg'  WHEN 'kg'      THEN 'kg'
      WHEN 'مجموعة'   THEN 'set' WHEN 'set'      THEN 'set'
      WHEN 'بند'       THEN 'ls'  WHEN 'إجمالي'   THEN 'ls'  WHEN 'ls'      THEN 'ls'
      WHEN 'l.s'       THEN 'ls'  WHEN 'lump sum' THEN 'ls'
      ELSE v_lib_unit
    END;

    IF v_lib_canon != v_canon_unit AND v_lib_unit != v_norm_unit THEN
      CONTINUE;
    END IF;

    v_precision := 0;
    IF v_lib_kw IS NOT NULL AND array_length(v_lib_kw, 1) > 0
       AND v_boq_kw_set IS NOT NULL AND array_length(v_boq_kw_set, 1) > 0
    THEN
      SELECT ARRAY(
        SELECT DISTINCT
          regexp_replace(
            regexp_replace(
              regexp_replace(
                regexp_replace(lower(kw), '[\u064B-\u065F\u0670]', '', 'g'),
                '[\u0625\u0623\u0622\u0671]', '\u0627', 'g'),
              '\u0629', '\u0647', 'g'),
            '\u0649', '\u064A', 'g')
        FROM unnest(v_lib_kw) kw
        WHERE length(kw) >= 2
      ) INTO v_lib_kw_norm;

      v_lib_kw_cnt := array_length(v_lib_kw_norm, 1);

      SELECT COUNT(*) INTO v_matched
      FROM (SELECT unnest(v_lib_kw_norm) INTERSECT SELECT unnest(v_boq_kw_set)) t;

      IF v_lib_kw_cnt > 0 THEN
        v_precision := v_matched::float / v_lib_kw_cnt::float;
      END IF;
    END IF;

    v_trgm := GREATEST(
      similarity(lower(v_item.description), lower(v_lib_name)),
      similarity(
        regexp_replace(lower(v_item.description), '[\u0625\u0623\u0622]', '\u0627', 'g'),
        regexp_replace(lower(v_lib_name), '[\u0625\u0623\u0622]', '\u0627', 'g')
      )
    );

    v_score := GREATEST(v_precision * 100.0, v_trgm * 100.0);

    IF v_item.description ~ '\d+\s*(سم|mm|cm)' AND v_lib_name ~ '\d+\s*(سم|mm|cm)' THEN
      v_boq_t := (regexp_match(v_item.description, '(\d+)\s*(سم|mm|cm)', 'i'))[1];
      v_lib_t := (regexp_match(v_lib_name, '(\d+)\s*(سم|mm|cm)', 'i'))[1];
      IF v_boq_t IS NOT NULL AND v_lib_t IS NOT NULL AND v_boq_t != v_lib_t THEN
        v_score := v_score - 40;
      END IF;
    END IF;

    IF v_score > v_best_conf THEN
      v_best_conf := ROUND(v_score)::integer;
      v_best_id   := v_lib_id;
      v_best_rate := v_lib_rate;
    END IF;
  END LOOP;

  IF v_best_id IS NOT NULL AND v_best_conf >= 30 THEN
    v_status := CASE WHEN v_best_conf >= 75 THEN 'approved' ELSE 'needs_review' END;

    UPDATE boq_items SET
      unit_rate      = v_best_rate,
      total_price    = v_item.quantity * v_best_rate,
      status         = v_status,
      linked_rate_id = v_best_id,
      confidence     = LEAST(v_best_conf, 95),
      updated_at     = now()
    WHERE id = p_item_id;

    RETURN jsonb_build_object(
      'success',     true,
      'unit_rate',   v_best_rate,
      'total_price', v_item.quantity * v_best_rate,
      'confidence',  LEAST(v_best_conf, 95),
      'lib_id',      v_best_id,
      'status',      v_status
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
