/*
  # BMS Calculations Table

  ## Purpose
  Stores BMS (Building Management System) point calculations per BOQ item.
  Each BMS item (identified by keywords like BMS/نظام إدارة المباني) gets
  priced at 500 SAR per BMS point, calculated from connected equipment.

  ## New Tables
  - `bms_calculations`
    - `id` (uuid, primary key)
    - `boq_item_id` (uuid, FK to boq_items) — the BMS line item being priced
    - `boq_file_id` (uuid, FK to boq_files) — for easy querying
    - `equipment_data` (jsonb) — snapshot of equipment counts entered by user
    - `total_points` (integer) — total BMS points calculated
    - `total_cost` (numeric) — total_points × 500
    - `price_per_point` (numeric) — always 500 SAR (stored for audit trail)
    - `created_by` (uuid)
    - `created_at` / `updated_at` (timestamptz)

  ## Security
  - RLS enabled
  - Authenticated users can manage their own calculations
*/

CREATE TABLE IF NOT EXISTS bms_calculations (
  id             uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  boq_item_id    uuid REFERENCES boq_items(id) ON DELETE CASCADE,
  boq_file_id    uuid REFERENCES boq_files(id) ON DELETE CASCADE,
  equipment_data jsonb NOT NULL DEFAULT '{}',
  total_points   integer NOT NULL DEFAULT 0,
  total_cost     numeric NOT NULL DEFAULT 0,
  price_per_point numeric NOT NULL DEFAULT 500,
  created_by     uuid REFERENCES auth.users(id),
  created_at     timestamptz DEFAULT now(),
  updated_at     timestamptz DEFAULT now()
);

ALTER TABLE bms_calculations ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users can select own bms_calculations"
  ON bms_calculations FOR SELECT
  TO authenticated
  USING (created_by = auth.uid());

CREATE POLICY "Users can insert own bms_calculations"
  ON bms_calculations FOR INSERT
  TO authenticated
  WITH CHECK (created_by = auth.uid());

CREATE POLICY "Users can update own bms_calculations"
  ON bms_calculations FOR UPDATE
  TO authenticated
  USING (created_by = auth.uid())
  WITH CHECK (created_by = auth.uid());

CREATE POLICY "Users can delete own bms_calculations"
  ON bms_calculations FOR DELETE
  TO authenticated
  USING (created_by = auth.uid());

CREATE INDEX IF NOT EXISTS idx_bms_calculations_boq_item ON bms_calculations(boq_item_id);
CREATE INDEX IF NOT EXISTS idx_bms_calculations_boq_file ON bms_calculations(boq_file_id);
