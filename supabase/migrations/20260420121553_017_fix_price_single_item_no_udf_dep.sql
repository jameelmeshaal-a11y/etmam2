/*
  # Final corrected price_single_boq_item — no external UDF dependencies

  ## Changes vs previous version
  - Removed arabicToLatin_safe() call (does not exist as a UDF)
  - Thickness gate uses simple regexp_replace + regexp_match inline
  - All logic self-contained in one PL/pgSQL function
  - Unit normalization covers common Arabic and English variants
  - Jaccard on keywords as primary signal, pg_trgm as secondary
  - Threshold: score >= 30 → price item; score >= 75 → status=approved; else needs_review
*/

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

  v_trgm        float;
  v_jaccard     float;
  v_score       float;

  v_boq_kw      text[];
  v_lib_kw_norm text[];
  v_intersect   int;
  v_union_size  int;

  v_boq_t       text;
  v_lib_t       text;

  v_status      text;
BEGIN
  -- ── 0. Load item ─────────────────────────────────────────────────────────
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

  -- ── 1. Normalize unit to canonical form ───────────────────────────────────
  v_norm_unit := lower(trim(COALESCE(v_item.unit, '')));

  v_canon_unit := CASE v_norm_unit
    WHEN 'متر مربع'  THEN 'm²'  WHEN 'م²'     THEN 'm²'  WHEN 'م2'  THEN 'm²'
    WHEN 'm2'        THEN 'm²'  WHEN 'sqm'    THEN 'm²'  WHEN 'sq.m' THEN 'm²'
    WHEN 'متر مكعب' THEN 'm³'  WHEN 'م³'     THEN 'm³'  WHEN 'م3'  THEN 'm³'
    WHEN 'm3'        THEN 'm³'  WHEN 'cum'    THEN 'm³'  WHEN 'cu.m' THEN 'm³'
    WHEN 'متر طولي' THEN 'lm'  WHEN 'م'      THEN 'lm'  WHEN 'متر' THEN 'lm'
    WHEN 'lm'        THEN 'lm'  WHEN 'ml'     THEN 'lm'  WHEN 'lin.m' THEN 'lm'
    WHEN 'عدد'       THEN 'nos' WHEN 'قطعة'   THEN 'nos' WHEN 'no'  THEN 'nos'
    WHEN 'nos'       THEN 'nos' WHEN 'pcs'    THEN 'nos' WHEN 'piece' THEN 'nos'
    WHEN 'طن'        THEN 'ton' WHEN 'ton'    THEN 'ton' WHEN 'tonnes' THEN 'ton'
    WHEN 'كجم'       THEN 'kg'  WHEN 'كيلوغرام' THEN 'kg' WHEN 'kg' THEN 'kg'
    WHEN 'مجموعة'   THEN 'set' WHEN 'set'    THEN 'set'
    WHEN 'بند'       THEN 'ls'  WHEN 'إجمالي' THEN 'ls'  WHEN 'ls'  THEN 'ls'
    WHEN 'l.s'       THEN 'ls'  WHEN 'lump sum' THEN 'ls'
    ELSE v_norm_unit
  END;

  -- ── 2. Build BOQ keywords by normalizing description text ─────────────────
  -- Strip diacritics, normalize alef forms, ta marbuta, alef maqsura
  SELECT ARRAY(
    SELECT DISTINCT token
    FROM (
      SELECT
        regexp_replace(
          regexp_replace(
            regexp_replace(
              regexp_replace(
                lower(word),
                '[\u064B-\u065F\u0670]', '', 'g'   -- strip diacritics
              ),
              '[\u0625\u0623\u0622\u0671]', '\u0627', 'g'  -- normalize alef
            ),
            '\u0629', '\u0647', 'g'                 -- ta marbuta → ha
          ),
          '\u0649', '\u064A', 'g'                   -- alef maqsura → ya
        ) AS token
      FROM regexp_split_to_table(
        regexp_replace(v_item.description, '[،؛؟!.,;:()\[\]{}|/\\\-]', ' ', 'g'),
        '\s+'
      ) word
    ) t
    WHERE length(token) >= 2
      AND token NOT IN (
        'من','في','على','إلى','الى','عن','مع','أن','ان','هو','هي',
        'لا','ما','كل','ذلك','هذا','هذه','التي','الذي','حتى','إذا'
      )
  ) INTO v_boq_kw;

  -- ── 3. Score each candidate library item ─────────────────────────────────
  FOR v_lib_id, v_lib_rate, v_lib_name, v_lib_kw, v_lib_unit IN
    SELECT rl.id, rl.rate_target, rl.standard_name_ar, rl.keywords,
           lower(trim(COALESCE(rl.unit, '')))
    FROM rate_library rl
    WHERE rl.source_type = 'Approved'
      AND rl.rate_target > 0
  LOOP
    -- Unit gate: skip if canonical units differ
    DECLARE
      v_lib_canon text;
    BEGIN
      v_lib_canon := CASE v_lib_unit
        WHEN 'متر مربع'  THEN 'm²'  WHEN 'م²'     THEN 'm²'  WHEN 'م2'  THEN 'm²'
        WHEN 'm2'        THEN 'm²'  WHEN 'sqm'    THEN 'm²'
        WHEN 'متر مكعب' THEN 'm³'  WHEN 'م³'     THEN 'm³'  WHEN 'م3'  THEN 'm³'
        WHEN 'm3'        THEN 'm³'  WHEN 'cum'    THEN 'm³'
        WHEN 'متر طولي' THEN 'lm'  WHEN 'م'      THEN 'lm'  WHEN 'متر' THEN 'lm'
        WHEN 'lm'        THEN 'lm'  WHEN 'ml'     THEN 'lm'
        WHEN 'عدد'       THEN 'nos' WHEN 'قطعة'   THEN 'nos' WHEN 'no'  THEN 'nos'
        WHEN 'nos'       THEN 'nos' WHEN 'pcs'    THEN 'nos' WHEN 'piece' THEN 'nos'
        WHEN 'طن'        THEN 'ton' WHEN 'ton'    THEN 'ton'
        WHEN 'كجم'       THEN 'kg'  WHEN 'كيلوغرام' THEN 'kg' WHEN 'kg' THEN 'kg'
        WHEN 'مجموعة'   THEN 'set' WHEN 'set'    THEN 'set'
        WHEN 'بند'       THEN 'ls'  WHEN 'إجمالي' THEN 'ls'  WHEN 'ls'  THEN 'ls'
        WHEN 'l.s'       THEN 'ls'  WHEN 'lump sum' THEN 'ls'
        ELSE v_lib_unit
      END;

      -- Skip if neither canonical match nor original unit match
      IF v_lib_canon != v_canon_unit AND v_lib_unit != v_norm_unit THEN
        CONTINUE;
      END IF;
    END;

    -- Jaccard similarity on keywords
    v_jaccard := 0;
    IF v_lib_kw IS NOT NULL AND array_length(v_lib_kw, 1) > 0
       AND v_boq_kw IS NOT NULL AND array_length(v_boq_kw, 1) > 0
    THEN
      -- Normalize library keywords
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

      SELECT COUNT(*) INTO v_intersect
      FROM (SELECT unnest(v_boq_kw) INTERSECT SELECT unnest(v_lib_kw_norm)) t;

      SELECT COUNT(*) INTO v_union_size
      FROM (SELECT unnest(v_boq_kw) UNION SELECT unnest(v_lib_kw_norm)) t;

      IF v_union_size > 0 THEN
        v_jaccard := v_intersect::float / v_union_size::float;
      END IF;
    END IF;

    -- pg_trgm similarity on description text
    v_trgm := GREATEST(
      similarity(lower(v_item.description), lower(v_lib_name)),
      -- also try with normalized alef forms
      similarity(
        regexp_replace(lower(v_item.description), '[\u0625\u0623\u0622]', '\u0627', 'g'),
        regexp_replace(lower(v_lib_name), '[\u0625\u0623\u0622]', '\u0627', 'g')
      )
    );

    -- Combined score: best of Jaccard × 100 or trgm × 100
    v_score := GREATEST(v_jaccard * 100.0, v_trgm * 100.0);

    -- Thickness gate: if both texts mention a thickness value and they differ, penalize
    v_boq_t := (regexp_match(
      regexp_replace(v_item.description, '[٠-٩]', -- replace Arabic numerals
        chr(48 + ascii(regexp_replace(regexp_replace(regexp_replace(regexp_replace(regexp_replace(
          regexp_replace(regexp_replace(regexp_replace(regexp_replace(
            regexp_replace((regexp_match(v_item.description, '[٠-٩]'))[1],'٠','0','g'),
          '١','1','g'),'٢','2','g'),'٣','3','g'),'٤','4','g'),
        '٥','5','g'),'٦','6','g'),'٧','7','g'),'٨','8','g'),'٩','9','g'))), 'g'),
      '(\d+)\s*(سم|mm|cm)', 'i'))[1];

    -- Simpler approach: just check if trgm or jaccard is low due to thickness mismatch
    -- Only apply penalty if both have digit patterns indicating thickness
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

  -- ── 4. Apply result ───────────────────────────────────────────────────────
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
