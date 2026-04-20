/*
  # Add unique constraint on bms_calculations.boq_item_id

  Required for upsert (ON CONFLICT boq_item_id) to work correctly.
  One BMS calculation per BOQ item.
*/

ALTER TABLE bms_calculations
  ADD CONSTRAINT bms_calculations_boq_item_id_unique UNIQUE (boq_item_id);
