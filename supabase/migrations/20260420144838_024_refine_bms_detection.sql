/*
  # Refine BMS Item Detection

  ## Problem
  The previous is_bms_item() function was too broad — it matched items like
  "مخرج اتصالات لزوم نظام BMS" (a data outlet FOR a BMS system) which should
  be auto-priced normally.

  ## Fix
  Tighten detection to only match items where BMS IS the main subject,
  not items that merely reference BMS as a context ("لزوم نظام BMS").
  
  The new logic requires BMS/نظام إدارة المباني to appear as the primary
  description, not as a suffix qualifier like "لزوم نظام BMS".
*/

CREATE OR REPLACE FUNCTION is_bms_item(p_description text)
RETURNS boolean
LANGUAGE plpgsql
IMMUTABLE
AS $$
DECLARE
  d text := trim(p_description);
  d_lower text := lower(trim(p_description));
BEGIN
  -- Must not be a "for BMS" suffix pattern (لزوم نظام BMS at end)
  IF d_lower ~ 'لزوم\s+نظام\s+bms\s*$' THEN
    RETURN false;
  END IF;
  IF d_lower ~ 'لزوم\s+.*bms\s*$' THEN
    RETURN false;
  END IF;

  -- BMS/BAS/DDC/SCADA as primary subject (not just mentioned in passing)
  -- Pattern: starts with or has BMS as a standalone word in main clause
  IF d ~ '^[^،,\n]*\(\s*BMS\s*\)' THEN RETURN true; END IF;  -- "نظام إدارة المباني ( BMS )"
  IF d ~ '^[^،,\n]*\(\s*BAS\s*\)' THEN RETURN true; END IF;
  IF d_lower ~ '^\s*(نظام\s+إدارة\s+المباني|نظام\s+ادارة\s+المباني)' THEN RETURN true; END IF;
  IF d_lower ~ '^\s*bms\b' THEN RETURN true; END IF;
  IF d_lower ~ '^\s*bas\b' THEN RETURN true; END IF;
  IF d_lower ~ '^\s*(building\s+management|building\s+automation)' THEN RETURN true; END IF;

  -- Multi-line descriptions where first line is a header and BMS is key subject
  IF d ~ E'\\n[^\n]*\(\\s*BMS\\s*\)' AND length(d) > 200 THEN RETURN true; END IF;

  -- DDC/SCADA as standalone primary system  
  IF d_lower ~ '^\s*(نظام\s+)?(ddc|scada)\b' THEN RETURN true; END IF;

  RETURN false;
END;
$$;

GRANT EXECUTE ON FUNCTION is_bms_item(text) TO authenticated;

-- Re-fix: restore the data outlet item to pending so it gets auto-priced
UPDATE boq_items
SET status = 'pending', unit_rate = NULL, total_price = NULL, 
    linked_rate_id = NULL, confidence = 0, updated_at = now()
WHERE is_bms_item(description) = false
  AND (
    lower(description) LIKE '%bms%'
    OR description LIKE '%نظام إدارة المباني%'
  )
  AND override_type IS NULL
  AND status = 'pending'
  AND unit_rate IS NULL;
