// ⛔ LOCKED FILE — DO NOT MODIFY WITHOUT EXPLICIT USER PERMISSION
// Last approved state: 2026-04-20
// Any change to this file requires user to say: "افتح ملف pricingEngine.ts"
import { supabase } from './supabase';
import type { RateLibraryItem } from '../types';

export interface PricingResult {
  itemId: string;
  success: boolean;
  skipped?: boolean;
  reason?: string;
  unitRate?: number;
  confidence?: number;
  libraryItemId?: string;
  error?: string;
}

export interface BatchPricingResult {
  total: number;
  priced: number;
  skipped: number;
  failed: number;
  results: PricingResult[];
}

export interface PricingProgress {
  current: number;
  total: number;
  pricedSoFar: number;
  failedSoFar: number;
  runningTotal: number;
  currentItem: string;
}

// ─── Item-by-item server-side pricing ────────────────────────────────────────
// Each item is priced with a separate RPC call to avoid statement timeouts.
// The onProgress callback is called after each item, giving real progress updates.

export async function priceItemsSequentially(
  boqFileId: string,
  _library: RateLibraryItem[],
  onProgress: (progress: PricingProgress) => void
): Promise<BatchPricingResult> {

  // 1. Load all items that need pricing from this file
  // Use a single SQL-level filter via RPC or raw query to avoid Supabase's
  // .not('col','eq','val') bug that also excludes NULLs (SQL: col != 'val' is NULL when col IS NULL)
  const { data: pendingItems, error: loadErr } = await supabase
    .from('boq_items')
    .select('id, description, quantity, status, override_type')
    .eq('boq_file_id', boqFileId)
    .in('status', ['pending', 'stale_price', 'needs_review'])
    .order('row_index', { ascending: true });

  if (loadErr) throw new Error(loadErr.message);

  const items = pendingItems ?? [];
  const total = items.length;

  if (total === 0) {
    onProgress({ current: 0, total: 0, pricedSoFar: 0, failedSoFar: 0, runningTotal: 0, currentItem: 'لا توجد بنود للتسعير' });
    return { total: 0, priced: 0, skipped: 0, failed: 0, results: [] };
  }

  let priced = 0;
  let failed = 0;
  let skipped = 0;
  let runningTotal = 0;
  const results: PricingResult[] = [];

  onProgress({ current: 0, total, pricedSoFar: 0, failedSoFar: 0, runningTotal: 0, currentItem: 'جاري التسعير...' });

  for (let i = 0; i < items.length; i++) {
    const item = items[i];

    // Report progress before processing this item
    onProgress({
      current: i,
      total,
      pricedSoFar: priced,
      failedSoFar: failed,
      runningTotal,
      currentItem: item.description?.substring(0, 60) ?? '',
    });

    // Call server-side RPC for this single item
    const { data, error } = await supabase.rpc('price_single_boq_item', {
      p_item_id: item.id,
    });

    if (error) {
      results.push({ itemId: item.id, success: false, error: error.message });
      failed++;
    } else {
      const r = data as {
        success: boolean;
        skipped?: boolean;
        reason?: string;
        unit_rate?: number;
        total_price?: number;
        confidence?: number;
        lib_id?: string;
      };

      if (r.skipped) {
        skipped++;
        results.push({ itemId: item.id, success: false, skipped: true, reason: r.reason });
      } else if (r.success) {
        priced++;
        runningTotal += r.total_price ?? 0;
        results.push({
          itemId: item.id,
          success: true,
          unitRate: r.unit_rate,
          confidence: r.confidence,
          libraryItemId: r.lib_id,
        });
      } else {
        failed++;
        results.push({ itemId: item.id, success: false, reason: r.reason, confidence: r.confidence });
      }
    }

    // Small yield to keep UI responsive
    await new Promise(resolve => setTimeout(resolve, 0));
  }

  // Final progress update
  onProgress({ current: total, total, pricedSoFar: priced, failedSoFar: failed, runningTotal, currentItem: 'اكتمل التسعير' });

  // Update file stats
  try {
    const { error: statsErr } = await supabase.rpc('price_boq_file_stats_only', { p_boq_file_id: boqFileId });
    if (statsErr) await updateFileStats(boqFileId);
  } catch {
    await updateFileStats(boqFileId).catch(console.error);
  }

  return { total, priced, skipped, failed, results };
}

// ─── Reprice a single item (called from row action button) ────────────────────

export async function repriceSingleItem(
  itemId: string,
  _library: RateLibraryItem[]
): Promise<PricingResult> {
  const { data, error } = await supabase.rpc('price_single_boq_item', {
    p_item_id: itemId,
  });

  if (error) return { itemId, success: false, error: error.message };

  const r = data as { success: boolean; skipped?: boolean; reason?: string; unit_rate?: number; confidence?: number; lib_id?: string };

  if (r.skipped) return { itemId, success: false, skipped: true, reason: r.reason };
  if (r.success) return { itemId, success: true, unitRate: r.unit_rate, confidence: r.confidence, libraryItemId: r.lib_id };
  return { itemId, success: false, reason: r.reason, confidence: r.confidence };
}

// ─── Update file stats ────────────────────────────────────────────────────────

export async function updateFileStats(boqFileId: string): Promise<void> {
  const { data: items } = await supabase
    .from('boq_items')
    .select('status, unit_rate, quantity')
    .eq('boq_file_id', boqFileId);

  if (!items) return;

  const priceable = items.filter(i => i.status !== 'descriptive' && (i.quantity ?? 0) > 0);
  const total = priceable.length;
  const priced = priceable.filter(i => i.unit_rate != null && i.unit_rate > 0).length;

  await supabase
    .from('boq_files')
    .update({ total_items: total, priced_items: priced })
    .eq('id', boqFileId);
}
