/*
  # Storage RLS Policies for BOQ Files

  Sets up Row Level Security for the boq-files storage bucket.
  Users can only access files they uploaded.
*/

CREATE POLICY "Users can upload BOQ files"
  ON storage.objects FOR INSERT
  TO authenticated
  WITH CHECK (bucket_id = 'boq-files' AND auth.uid()::text = (storage.foldername(name))[1]);

CREATE POLICY "Users can view their BOQ files"
  ON storage.objects FOR SELECT
  TO authenticated
  USING (bucket_id = 'boq-files' AND auth.uid()::text = (storage.foldername(name))[1]);

CREATE POLICY "Users can delete their BOQ files"
  ON storage.objects FOR DELETE
  TO authenticated
  USING (bucket_id = 'boq-files' AND auth.uid()::text = (storage.foldername(name))[1]);
