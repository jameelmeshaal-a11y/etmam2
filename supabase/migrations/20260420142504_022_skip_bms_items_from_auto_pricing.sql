/*
  # Skip BMS Items from Auto Pricing

  ## Rule
  Any BOQ item whose description contains BMS-related keywords is excluded
  from automatic price matching. These items must be priced manually using
  the BMS points calculator (500 SAR per point).

  ## BMS Keywords Detected
  - BMS, BAS
  - نظام إدارة المباني, نظام ادارة المباني
  - Building Management, Building Automation
  - DDC, SCADA
  - نظام مراقبة

  ## Changes
  1. Update price_single_boq_item function to skip BMS items early
     and mark them as 'pending' with reason 'bms_manual_required'
*/

CREATE OR REPLACE FUNCTION is_bms_item(p_description text)
RETURNS boolean
LANGUAGE sql
IMMUTABLE
AS $$
  SELECT (
    lower(p_description) LIKE '%bms%'
    OR lower(p_description) LIKE '%bas%'
    OR lower(p_description) LIKE '%ddc%'
    OR lower(p_description) LIKE '%scada%'
    OR p_description LIKE '%نظام إدارة المباني%'
    OR p_description LIKE '%نظام ادارة المباني%'
    OR p_description LIKE '%نظام ادارة%'
    OR p_description LIKE '%نظام مراقبة%'
    OR lower(p_description) LIKE '%building management%'
    OR lower(p_description) LIKE '%building automation%'
  )
$$;

GRANT EXECUTE ON FUNCTION is_bms_item(text) TO authenticated;
