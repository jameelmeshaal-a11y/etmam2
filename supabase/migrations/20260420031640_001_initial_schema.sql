/*
  # AI Construction Cost Estimation Engine — Initial Schema

  ## Overview
  This migration creates the full database schema for the BOQ pricing engine
  for Saudi Arabia construction projects (Etimad platform integration).

  ## New Tables
  1. `projects` — Project metadata (name, client, city)
  2. `boq_files` — Uploaded Excel file metadata per project
  3. `rate_library` — The single source of truth for all unit rates
  4. `boq_items` — Per-line pricing state for each BOQ item
  5. `rate_sources` — Audit trail for every price change

  ## Security
  - RLS enabled on all tables
  - Authenticated users can manage their own data
  - rate_library readable by all authenticated users
  - Governance: is_locked records protected via RLS

  ## Notes
  1. rate_library.keywords and item_name_aliases are auto-generated token arrays
  2. boq_items.override_type = 'manual' permanently locks item from repricing
  3. boq_items.status: pending / approved / stale_price / descriptive / needs_review
  4. rate_library.source_type: Approved / Field-Approved / Draft
*/

-- ============================================================
-- PROJECTS
-- ============================================================
CREATE TABLE IF NOT EXISTS projects (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  name text NOT NULL,
  client text DEFAULT '',
  city text DEFAULT '',
  created_by uuid REFERENCES auth.users(id) ON DELETE SET NULL,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

ALTER TABLE projects ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users can view own projects"
  ON projects FOR SELECT
  TO authenticated
  USING (created_by = auth.uid());

CREATE POLICY "Users can insert own projects"
  ON projects FOR INSERT
  TO authenticated
  WITH CHECK (created_by = auth.uid());

CREATE POLICY "Users can update own projects"
  ON projects FOR UPDATE
  TO authenticated
  USING (created_by = auth.uid())
  WITH CHECK (created_by = auth.uid());

CREATE POLICY "Users can delete own projects"
  ON projects FOR DELETE
  TO authenticated
  USING (created_by = auth.uid());

-- ============================================================
-- BOQ_FILES
-- ============================================================
CREATE TABLE IF NOT EXISTS boq_files (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  project_id uuid REFERENCES projects(id) ON DELETE CASCADE,
  name text NOT NULL,
  storage_path text NOT NULL DEFAULT '',
  city text DEFAULT '',
  created_at timestamptz DEFAULT now(),
  total_items integer DEFAULT 0,
  priced_items integer DEFAULT 0,
  export_variance_pct numeric DEFAULT 0,
  created_by uuid REFERENCES auth.users(id) ON DELETE SET NULL
);

ALTER TABLE boq_files ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users can view own boq files"
  ON boq_files FOR SELECT
  TO authenticated
  USING (created_by = auth.uid());

CREATE POLICY "Users can insert own boq files"
  ON boq_files FOR INSERT
  TO authenticated
  WITH CHECK (created_by = auth.uid());

CREATE POLICY "Users can update own boq files"
  ON boq_files FOR UPDATE
  TO authenticated
  USING (created_by = auth.uid())
  WITH CHECK (created_by = auth.uid());

CREATE POLICY "Users can delete own boq files"
  ON boq_files FOR DELETE
  TO authenticated
  USING (created_by = auth.uid());

-- ============================================================
-- RATE_LIBRARY — The Price Authority
-- ============================================================
CREATE TABLE IF NOT EXISTS rate_library (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  standard_name_ar text NOT NULL,
  standard_name_en text DEFAULT '',
  category text NOT NULL DEFAULT 'general',
  unit text NOT NULL,
  rate_base numeric NOT NULL DEFAULT 0,
  rate_target numeric NOT NULL DEFAULT 0,
  rate_min numeric NOT NULL DEFAULT 0,
  rate_max numeric NOT NULL DEFAULT 0,
  keywords text[] DEFAULT '{}',
  item_name_aliases text[] DEFAULT '{}',
  is_locked boolean DEFAULT false,
  source_type text NOT NULL DEFAULT 'Draft' CHECK (source_type IN ('Approved', 'Field-Approved', 'Draft')),
  approved_at timestamptz,
  last_reviewed_at timestamptz DEFAULT now(),
  created_by uuid REFERENCES auth.users(id) ON DELETE SET NULL,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_rate_library_category ON rate_library(category);
CREATE INDEX IF NOT EXISTS idx_rate_library_unit ON rate_library(unit);
CREATE INDEX IF NOT EXISTS idx_rate_library_source_type ON rate_library(source_type);
CREATE INDEX IF NOT EXISTS idx_rate_library_keywords ON rate_library USING GIN(keywords);
CREATE INDEX IF NOT EXISTS idx_rate_library_aliases ON rate_library USING GIN(item_name_aliases);

ALTER TABLE rate_library ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Authenticated users can view rate library"
  ON rate_library FOR SELECT
  TO authenticated
  USING (true);

CREATE POLICY "Users can insert rate library records"
  ON rate_library FOR INSERT
  TO authenticated
  WITH CHECK (true);

CREATE POLICY "Users can update unlocked rate library records"
  ON rate_library FOR UPDATE
  TO authenticated
  USING (is_locked = false)
  WITH CHECK (is_locked = false);

CREATE POLICY "Users can delete unlocked non-approved rate library records"
  ON rate_library FOR DELETE
  TO authenticated
  USING (is_locked = false AND source_type <> 'Approved');

-- ============================================================
-- BOQ_ITEMS — Per-Line Pricing State
-- ============================================================
CREATE TABLE IF NOT EXISTS boq_items (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  boq_file_id uuid REFERENCES boq_files(id) ON DELETE CASCADE,
  item_no text DEFAULT '',
  description text DEFAULT '',
  unit text DEFAULT '',
  quantity numeric DEFAULT 0,
  unit_rate numeric,
  total_price numeric,
  status text NOT NULL DEFAULT 'pending' CHECK (status IN ('pending', 'approved', 'stale_price', 'descriptive', 'needs_review', 'manual')),
  override_type text CHECK (override_type IN ('manual', NULL)),
  linked_rate_id uuid,
  confidence integer DEFAULT 0 CHECK (confidence >= 0 AND confidence <= 100),
  row_index integer DEFAULT 0,
  materials numeric DEFAULT 0,
  labor numeric DEFAULT 0,
  equipment numeric DEFAULT 0,
  logistics numeric DEFAULT 0,
  risk numeric DEFAULT 0,
  profit numeric DEFAULT 0,
  created_at timestamptz DEFAULT now(),
  updated_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_boq_items_boq_file_id ON boq_items(boq_file_id);
CREATE INDEX IF NOT EXISTS idx_boq_items_status ON boq_items(status);
CREATE INDEX IF NOT EXISTS idx_boq_items_override_type ON boq_items(override_type);
CREATE INDEX IF NOT EXISTS idx_boq_items_linked_rate_id ON boq_items(linked_rate_id);

ALTER TABLE boq_items ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users can view boq items for their files"
  ON boq_items FOR SELECT
  TO authenticated
  USING (
    EXISTS (
      SELECT 1 FROM boq_files bf
      WHERE bf.id = boq_items.boq_file_id
      AND bf.created_by = auth.uid()
    )
  );

CREATE POLICY "Users can insert boq items for their files"
  ON boq_items FOR INSERT
  TO authenticated
  WITH CHECK (
    EXISTS (
      SELECT 1 FROM boq_files bf
      WHERE bf.id = boq_items.boq_file_id
      AND bf.created_by = auth.uid()
    )
  );

CREATE POLICY "Users can update boq items for their files"
  ON boq_items FOR UPDATE
  TO authenticated
  USING (
    EXISTS (
      SELECT 1 FROM boq_files bf
      WHERE bf.id = boq_items.boq_file_id
      AND bf.created_by = auth.uid()
    )
  )
  WITH CHECK (
    EXISTS (
      SELECT 1 FROM boq_files bf
      WHERE bf.id = boq_items.boq_file_id
      AND bf.created_by = auth.uid()
    )
  );

CREATE POLICY "Users can delete boq items for their files"
  ON boq_items FOR DELETE
  TO authenticated
  USING (
    EXISTS (
      SELECT 1 FROM boq_files bf
      WHERE bf.id = boq_items.boq_file_id
      AND bf.created_by = auth.uid()
    )
  );

-- ============================================================
-- RATE_SOURCES — Audit Trail
-- ============================================================
CREATE TABLE IF NOT EXISTS rate_sources (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  boq_item_id uuid REFERENCES boq_items(id) ON DELETE CASCADE,
  rate_library_id uuid,
  unit_rate numeric NOT NULL,
  source_type text NOT NULL DEFAULT 'auto',
  override_type text,
  materials numeric DEFAULT 0,
  labor numeric DEFAULT 0,
  equipment numeric DEFAULT 0,
  logistics numeric DEFAULT 0,
  risk numeric DEFAULT 0,
  profit numeric DEFAULT 0,
  created_by uuid REFERENCES auth.users(id) ON DELETE SET NULL,
  created_at timestamptz DEFAULT now()
);

CREATE INDEX IF NOT EXISTS idx_rate_sources_boq_item_id ON rate_sources(boq_item_id);

ALTER TABLE rate_sources ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Users can view rate sources for their items"
  ON rate_sources FOR SELECT
  TO authenticated
  USING (
    EXISTS (
      SELECT 1 FROM boq_items bi
      JOIN boq_files bf ON bf.id = bi.boq_file_id
      WHERE bi.id = rate_sources.boq_item_id
      AND bf.created_by = auth.uid()
    )
  );

CREATE POLICY "Users can insert rate sources for their items"
  ON rate_sources FOR INSERT
  TO authenticated
  WITH CHECK (
    EXISTS (
      SELECT 1 FROM boq_items bi
      JOIN boq_files bf ON bf.id = bi.boq_file_id
      WHERE bi.id = rate_sources.boq_item_id
      AND bf.created_by = auth.uid()
    )
  );
