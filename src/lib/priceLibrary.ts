// ⛔ LOCKED FILE — DO NOT MODIFY WITHOUT EXPLICIT USER PERMISSION
// Last approved state: 2026-04-20
// Any change to this file requires user to say: "افتح ملف priceLibrary.ts"
import { supabase } from './supabase';
import { generateKeywords } from './matchingV3';
import type { RateLibraryItem } from '../types';

const FETCH_PAGE_SIZE = 1000;

async function fetchAllPages(query: () => ReturnType<typeof supabase.from>): Promise<RateLibraryItem[]> {
  const allItems: RateLibraryItem[] = [];
  let from = 0;

  while (true) {
    const { data, error } = await supabase
      .from('rate_library')
      .select('*')
      .order('category', { ascending: true })
      .order('standard_name_ar', { ascending: true })
      .range(from, from + FETCH_PAGE_SIZE - 1);

    if (error) throw error;
    if (!data || data.length === 0) break;

    allItems.push(...(data as RateLibraryItem[]));
    if (data.length < FETCH_PAGE_SIZE) break;
    from += FETCH_PAGE_SIZE;
  }

  return allItems;
}

export async function getAllLibraryItems(): Promise<RateLibraryItem[]> {
  return fetchAllPages(() => supabase.from('rate_library'));
}

export async function getApprovedLibraryItems(): Promise<RateLibraryItem[]> {
  const allItems: RateLibraryItem[] = [];
  let from = 0;

  while (true) {
    const { data, error } = await supabase
      .from('rate_library')
      .select('*')
      .eq('source_type', 'Approved')
      .order('category', { ascending: true })
      .range(from, from + FETCH_PAGE_SIZE - 1);

    if (error) throw error;
    if (!data || data.length === 0) break;

    allItems.push(...(data as RateLibraryItem[]));
    if (data.length < FETCH_PAGE_SIZE) break;
    from += FETCH_PAGE_SIZE;
  }

  return allItems;
}

export async function upsertLibraryItem(
  item: Partial<RateLibraryItem> & { standard_name_ar: string; unit: string; category: string }
): Promise<RateLibraryItem> {
  const keywords = generateKeywords(item.standard_name_ar);
  const aliases = [...new Set([
    ...(item.item_name_aliases ?? []),
    item.standard_name_ar,
    ...(item.standard_name_en ? [item.standard_name_en] : []),
  ])];

  const payload = {
    ...item,
    keywords,
    item_name_aliases: aliases,
    updated_at: new Date().toISOString(),
    ...(item.source_type === 'Approved' && !item.approved_at
      ? { approved_at: new Date().toISOString() }
      : {}),
  };

  if (item.id) {
    const { data: existing } = await supabase
      .from('rate_library')
      .select('is_locked')
      .eq('id', item.id)
      .maybeSingle();

    if (existing?.is_locked) {
      throw new Error(`Cannot edit locked record: ${item.standard_name_ar}`);
    }

    const { data, error } = await supabase
      .from('rate_library')
      .update(payload)
      .eq('id', item.id)
      .select()
      .single();

    if (error) throw error;
    return data as RateLibraryItem;
  }

  const { data: currentUser } = await supabase.auth.getUser();
  const { data, error } = await supabase
    .from('rate_library')
    .insert({ ...payload, created_by: currentUser.user?.id })
    .select()
    .single();

  if (error) throw error;
  return data as RateLibraryItem;
}

export async function deleteLibraryItem(id: string): Promise<void> {
  const { data: item } = await supabase
    .from('rate_library')
    .select('is_locked, source_type')
    .eq('id', id)
    .maybeSingle();

  if (item?.is_locked) {
    throw new Error('Cannot delete a locked record.');
  }
  if (item?.source_type === 'Approved') {
    throw new Error('Cannot delete an Approved record.');
  }

  const { error } = await supabase
    .from('rate_library')
    .delete()
    .eq('id', id);

  if (error) throw error;
}

export async function lockLibraryItem(id: string, lock: boolean): Promise<void> {
  const { error } = await supabase
    .from('rate_library')
    .update({ is_locked: lock, updated_at: new Date().toISOString() })
    .eq('id', id);
  if (error) throw error;
}

const BATCH_SIZE = 500;

export async function bulkImportLibraryItems(
  items: Array<Partial<RateLibraryItem> & { standard_name_ar: string; unit: string; category: string }>,
  onProgress?: (done: number, total: number) => void
): Promise<{ success: number; failed: number; errors: string[] }> {
  let success = 0;
  let failed = 0;
  const errors: string[] = [];
  const { data: currentUser } = await supabase.auth.getUser();
  const now = new Date().toISOString();

  const batches: typeof items[] = [];
  for (let i = 0; i < items.length; i += BATCH_SIZE) {
    batches.push(items.slice(i, i + BATCH_SIZE));
  }

  let done = 0;
  for (const batch of batches) {
    const rows = batch.map(item => {
      const keywords = generateKeywords(item.standard_name_ar);
      const aliases = [...new Set([
        ...(item.item_name_aliases ?? []),
        item.standard_name_ar,
        ...(item.standard_name_en ? [item.standard_name_en] : []),
      ])];
      return {
        standard_name_ar: item.standard_name_ar,
        standard_name_en: item.standard_name_en ?? '',
        category: item.category,
        unit: item.unit,
        rate_base: item.rate_base ?? 0,
        rate_target: item.rate_target ?? item.rate_base ?? 0,
        rate_min: item.rate_min ?? 0,
        rate_max: item.rate_max ?? 0,
        source_type: item.source_type ?? 'Approved',
        keywords,
        item_name_aliases: aliases,
        is_locked: false,
        approved_at: item.source_type === 'Approved' ? now : null,
        created_by: currentUser.user?.id ?? null,
        created_at: now,
        updated_at: now,
      };
    });

    const { error } = await supabase.from('rate_library').insert(rows);
    if (error) {
      failed += batch.length;
      errors.push(error.message);
    } else {
      success += batch.length;
    }
    done += batch.length;
    onProgress?.(done, items.length);
  }

  return { success, failed, errors };
}

export async function lockAllLibraryItems(lock: boolean): Promise<void> {
  const { error } = await supabase
    .from('rate_library')
    .update({ is_locked: lock, updated_at: new Date().toISOString() })
    .neq('id', '00000000-0000-0000-0000-000000000000');
  if (error) throw error;
}
