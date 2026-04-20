/*
  # Improved price_single_boq_item with multi-stage matching

  ## Problem
  The previous version used only pg_trgm similarity which performs poorly on Arabic text.
  It also ignored the keywords field and didn't normalize Arabic units, causing items
  to remain unpriced even when matching library entries exist.

  ## Solution
  New 4-stage matching algorithm in SQL mirroring matchingV3.ts:

  Stage 1 — Unit normalization: Map Arabic units to canonical form (متر مربع → m²)
  Stage 2 — Keyword Jaccard similarity: Compare BOQ keywords vs library keywords
  Stage 3 — Word/description similarity: pg_trgm on normalized description vs standard_name_ar
  Stage 4 — Combined score: max(Jaccard * 100, trgm * 100) with thickness gate

  Threshold: >= 30 to price, >= 75 for status=approved, else status=needs_review

  ## Changes
  - Recreates price_single_boq_item with improved matching
  - No schema changes, no data loss
*/

CREATE OR REPLACE FUNCTION price_single_boq_item(p_item_id uuid)
RETURNS jsonb
LANGUAGE plpgsql
SECURITY DEFINER
AS $$
DECLARE
  v_item        RECORD;
  v_norm_unit   text;
  v_best_id     uuid;
  v_best_rate   numeric;
  v_best_conf   integer;
  v_best_status text;

  v_lib_id      uuid;
  v_lib_rate    numeric;
  v_lib_name    text;
  v_lib_kw      text[];
  v_lib_unit    text;

  v_trgm_ar     float;
  v_trgm_en     float;
  v_trgm        float;
  v_jaccard     float;
  v_score       float;

  v_boq_kw      text[];
  v_boq_kw_norm text[];
  v_lib_kw_norm text[];
  v_intersect   int;
  v_union_size  int;

  v_boq_thick   int;
  v_lib_thick   int;
  v_thick_match text;

  -- Unit normalization map keys (Arabic → canonical)
  v_unit_map_key   text[];
  v_unit_map_val   text[];
  v_canon_unit     text;
  i                int;
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

  -- ── 1. Normalize unit (Arabic → canonical) ────────────────────────────────
  v_unit_map_key := ARRAY[
    'متر مربع','م²','م2','m2','sqm','sq.m',
    'متر مكعب','م³','م3','m3','cum','cu.m',
    'متر طولي','م','متر','lm','ml','lin.m','linear m',
    'عدد','قطعة','no','nos.','no.','pcs','piece','pieces',
    'طن','ton.','tonnes',
    'كجم','كيلوغرام','kg.','kilogram',
    'مجموعة','set.','sets',
    'بند','إجمالي','l.s','lump sum','ls.'
  ];
  v_unit_map_val := ARRAY[
    'm²','m²','m²','m²','m²','m²',
    'm³','m³','m³','m³','m³','m³',
    'lm','lm','lm','lm','lm','lm','lm',
    'nos','nos','nos','nos','nos','nos','nos','nos',
    'ton','ton','ton',
    'kg','kg','kg','kg',
    'set','set','set',
    'ls','ls','ls','ls','ls'
  ];

  v_norm_unit := lower(trim(v_item.unit));
  v_canon_unit := v_norm_unit;
  FOR i IN 1..array_length(v_unit_map_key, 1) LOOP
    IF v_norm_unit = lower(trim(v_unit_map_key[i])) THEN
      v_canon_unit := v_unit_map_val[i];
      EXIT;
    END IF;
  END LOOP;

  -- ── 2. Build BOQ keywords from description ────────────────────────────────
  -- Normalize Arabic: remove diacritics, normalize alef, ta marbuta, alef maqsura
  -- Split by whitespace and punctuation, remove short tokens
  SELECT ARRAY(
    SELECT DISTINCT token
    FROM (
      SELECT regexp_replace(
               regexp_replace(
                 regexp_replace(
                   regexp_replace(lower(unnest(string_to_array(
                     regexp_replace(v_item.description, '[\u060C\u061B\u061F\u0021-\u002F\u003A-\u0040\u005B-\u0060\u007B-\u007E،؟!()[\]{}/\\|.,;:\-]', ' ', 'g'),
                     ' '
                   ))), '[\u064B-\u065F\u0670]', '', 'g'),  -- strip diacritics
                 '[\u0625\u0623\u0622]', '\u0627', 'g'),  -- normalize alef
               '\u0629', '\u0647', 'g'),                  -- ta marbuta → ha
             '\u0649', '\u064A', 'g') AS token            -- alef maqsura → ya
    ) t
    WHERE length(token) >= 2
      AND token NOT IN ('من','في','على','إلى','الى','عن','مع','أن','ان','هو','هي','لا','ما','كل')
  ) INTO v_boq_kw;

  -- ── 3. Find best matching library item ───────────────────────────────────
  v_best_id   := NULL;
  v_best_rate := NULL;
  v_best_conf := 0;

  FOR v_lib_id, v_lib_rate, v_lib_name, v_lib_kw, v_lib_unit IN
    SELECT rl.id, rl.rate_target, rl.standard_name_ar, rl.keywords, lower(trim(rl.unit))
    FROM rate_library rl
    WHERE rl.source_type = 'Approved'
      AND rl.rate_target > 0
      AND (
        -- Exact canonical unit match
        lower(trim(rl.unit)) = v_canon_unit
        -- Original unit match
        OR lower(trim(rl.unit)) = v_norm_unit
        -- Library unit normalizes to same canonical
        OR (
          CASE lower(trim(rl.unit))
            WHEN 'متر مربع' THEN 'm²' WHEN 'م²' THEN 'm²' WHEN 'م2' THEN 'm²' WHEN 'm2' THEN 'm²' WHEN 'sqm' THEN 'm²'
            WHEN 'متر مكعب' THEN 'm³' WHEN 'م³' THEN 'm³' WHEN 'م3' THEN 'm³' WHEN 'm3' THEN 'm³' WHEN 'cum' THEN 'm³'
            WHEN 'متر طولي' THEN 'lm' WHEN 'م' THEN 'lm' WHEN 'متر' THEN 'lm' WHEN 'ml' THEN 'lm'
            WHEN 'عدد' THEN 'nos' WHEN 'قطعة' THEN 'nos' WHEN 'no' THEN 'nos' WHEN 'pcs' THEN 'nos' WHEN 'piece' THEN 'nos'
            WHEN 'طن' THEN 'ton'
            WHEN 'كجم' THEN 'kg' WHEN 'كيلوغرام' THEN 'kg'
            WHEN 'مجموعة' THEN 'set'
            WHEN 'بند' THEN 'ls' WHEN 'إجمالي' THEN 'ls' WHEN 'l.s' THEN 'ls' WHEN 'lump sum' THEN 'ls'
            ELSE lower(trim(rl.unit))
          END
        ) = v_canon_unit
      )
  LOOP
    -- ── Jaccard on keywords ──────────────────────────────────────────────
    v_jaccard := 0;
    IF v_lib_kw IS NOT NULL AND array_length(v_lib_kw, 1) > 0
       AND v_boq_kw IS NOT NULL AND array_length(v_boq_kw, 1) > 0 THEN

      -- Normalize library keywords same way
      SELECT ARRAY(
        SELECT DISTINCT regexp_replace(
          regexp_replace(
            regexp_replace(
              regexp_replace(lower(kw), '[\u064B-\u065F\u0670]', '', 'g'),
            '[\u0625\u0623\u0622]', '\u0627', 'g'),
          '\u0629', '\u0647', 'g'),
        '\u0649', '\u064A', 'g')
        FROM unnest(v_lib_kw) kw
        WHERE length(kw) >= 2
      ) INTO v_lib_kw_norm;

      -- intersection count
      SELECT COUNT(*) INTO v_intersect
      FROM (
        SELECT unnest(v_boq_kw)
        INTERSECT
        SELECT unnest(v_lib_kw_norm)
      ) t;

      -- union count
      SELECT COUNT(*) INTO v_union_size
      FROM (
        SELECT unnest(v_boq_kw)
        UNION
        SELECT unnest(v_lib_kw_norm)
      ) t;

      IF v_union_size > 0 THEN
        v_jaccard := v_intersect::float / v_union_size::float;
      END IF;
    END IF;

    -- ── pg_trgm similarity on description ────────────────────────────────
    v_trgm_ar := similarity(lower(v_item.description), lower(v_lib_name));
    v_trgm_en := 0;
    v_trgm    := v_trgm_ar;

    -- ── Combined score: best of Jaccard or trgm ──────────────────────────
    v_score := GREATEST(v_jaccard * 100, v_trgm * 100);

    -- ── Thickness gate: penalize -40 if thicknesses disagree ─────────────
    -- Extract first integer from description (simple heuristic)
    IF v_item.description ~ '\d+\s*(سم|mm|cm)' AND v_lib_name ~ '\d+\s*(سم|mm|cm)' THEN
      DECLARE
        v_boq_t text := (regexp_match(arabicToLatin_safe(v_item.description), '(\d+)\s*(سم|mm|cm)'))[1];
        v_lib_t text := (regexp_match(arabicToLatin_safe(v_lib_name), '(\d+)\s*(سم|mm|cm)'))[1];
      BEGIN
        IF v_boq_t IS NOT NULL AND v_lib_t IS NOT NULL AND v_boq_t != v_lib_t THEN
          v_score := v_score - 40;
        END IF;
      EXCEPTION WHEN OTHERS THEN
        NULL; -- ignore errors in thickness gate
      END;
    END IF;

    IF v_score > v_best_conf THEN
      v_best_conf := ROUND(v_score)::integer;
      v_best_id   := v_lib_id;
      v_best_rate := v_lib_rate;
    END IF;
  END LOOP;

  -- ── 4. Apply result ───────────────────────────────────────────────────────
  IF v_best_id IS NOT NULL AND v_best_conf >= 30 THEN
    v_best_status := CASE WHEN v_best_conf >= 75 THEN 'approved' ELSE 'needs_review' END;

    UPDATE boq_items SET
      unit_rate      = v_best_rate,
      total_price    = v_item.quantity * v_best_rate,
      status         = v_best_status,
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
      'status',      v_best_status
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
