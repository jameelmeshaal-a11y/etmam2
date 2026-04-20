/*
  # Fix Zero-Quantity Items as Descriptive

  ## Purpose
  Items with quantity = 0 or NULL that have no unit rate should be treated as
  descriptive/header rows, not as unpriced items needing attention.

  ## Changes
  - Updates boq_items where quantity = 0 or null AND unit_rate is null AND
    status is not already 'descriptive' → sets status to 'descriptive'
  - Also updates price_single_boq_item RPC to skip qty=0 items (mark as descriptive)

  ## Notes
  - Only affects items without a unit_rate (truly unpriced structural rows)
  - Does not touch items that have a price but qty=0
*/

UPDATE boq_items
SET status = 'descriptive', updated_at = now()
WHERE (quantity IS NULL OR quantity = 0)
  AND unit_rate IS NULL
  AND status != 'descriptive';
