import type { BOQItem, RateLibraryItem, MatchResult } from '../types';

// ─── Arabic text normalisation ───────────────────────────────────────────────

export function normalizeArabic(text: string): string {
  return text
    .replace(/[\u064B-\u065F\u0670]/g, '')
    .replace(/[\u0625\u0623\u0622]/g, '\u0627')
    .replace(/\u0629/g, '\u0647')
    .replace(/\u0649/g, '\u064A')
    .trim()
    .toLowerCase();
}

export function generateKeywords(name: string): string[] {
  const stripped = normalizeArabic(name);
  const stopPfx = /^(\u0627\u0644|\u0648\u0627\u0644|\u0628\u0627\u0644|\u0648|\u0628|\u0644)/;
  return [...new Set(
    stripped
      .split(/[\s,\u060C.;/\\()\-\u2013]+/)
      .map(t => t.replace(stopPfx, ''))
      .filter(t => t.length >= 2)
  )];
}

// ─── Item number normalisation ────────────────────────────────────────────────

export function normalizeItemNo(raw: string): string {
  return String(raw ?? '')
    .replace(/\u00A0/g, '')
    .replace(/\u200B/g, '')
    .replace(/[\u0660-\u0669]/g, d => String(d.charCodeAt(0) - 0x0660))
    .replace(/\s+/g, '')
    .trim()
    .toLowerCase();
}

// ─── Unit normalisation ───────────────────────────────────────────────────────

const UNIT_MAP: Record<string, string> = {
  'متر مربع': 'm²', 'م²': 'm²', 'م2': 'm²',
  'متر مكعب': 'm³', 'م³': 'm³', 'م3': 'm³',
  'متر طولي': 'lm', 'م': 'lm', 'متر': 'lm',
  'عدد': 'nos', 'قطعة': 'nos',
  'طن': 'ton', 'كجم': 'kg', 'كيلوغرام': 'kg',
  'مجموعة': 'set', 'بند': 'ls', 'إجمالي': 'ls',
};

export function normalizeUnit(unit: string): string {
  const u = unit.trim().toLowerCase();
  return UNIT_MAP[u] ?? u;
}

// ─── Similarity helpers ───────────────────────────────────────────────────────

function jaccardSimilarity(a: string[], b: string[]): number {
  if (!a.length || !b.length) return 0;
  const setA = new Set(a);
  const setB = new Set(b);
  const intersection = [...setA].filter(x => setB.has(x)).length;
  const union = new Set([...setA, ...setB]).size;
  return union === 0 ? 0 : intersection / union;
}

function wordSimilarity(a: string, b: string): number {
  const na = normalizeArabic(a);
  const nb = normalizeArabic(b);
  if (na === nb) return 1;
  if (na.includes(nb) || nb.includes(na)) return 0.85;
  const wa = na.split(/\s+/);
  const wb = nb.split(/\s+/);
  return jaccardSimilarity(wa, wb);
}

// ─── Gate A — Thickness Gate ──────────────────────────────────────────────────

const ARABIC_DIGIT_MAP: Record<string, string> = {
  '٠': '0', '١': '1', '٢': '2', '٣': '3', '٤': '4',
  '٥': '5', '٦': '6', '٧': '7', '٨': '8', '٩': '9',
};

function arabicToLatin(s: string): string {
  return s.replace(/[٠-٩]/g, d => ARABIC_DIGIT_MAP[d] ?? d);
}

function extractThickness(text: string): number | null {
  const t = arabicToLatin(text);
  const patterns = [
    /(\d+)\s*mm/i,
    /(\d+)\s*سم/,
    /(\d+)\s*ملم/,
    /سماكة\s*(\d+)/,
    /thickness\s*(\d+)/i,
    /(\d+)\s*cm/i,
  ];
  for (const p of patterns) {
    const m = t.match(p);
    if (m) {
      const val = parseInt(m[1]);
      if (p.source.includes('cm') || p.source.includes('سم')) return val * 10;
      return val;
    }
  }
  return null;
}

function applyThicknessGate(boqText: string, libText: string, score: number): number {
  const boqThick = extractThickness(boqText);
  const libThick = extractThickness(libText);
  if (boqThick !== null && libThick !== null && boqThick !== libThick) {
    return score - 40;
  }
  return score;
}

// ─── Gate B — Fire Rating Gate ────────────────────────────────────────────────

function isFireRated(text: string): boolean {
  return /fire.?rat|مقاوم.*حريق|حريق.*مقاوم|fire resist/i.test(text);
}

function extractFireMinutes(text: string): number | null {
  const m = arabicToLatin(text).match(/(\d+)\s*(?:min|دقيقة|دق)/i);
  if (m) return parseInt(m[1]);
  if (/60/.test(text)) return 60;
  if (/90/.test(text)) return 90;
  if (/120/.test(text)) return 120;
  if (/180/.test(text)) return 180;
  return null;
}

function applyFireRatingGate(boqText: string, libText: string): boolean {
  const boqFire = isFireRated(boqText);
  const libFire = isFireRated(libText);
  if (boqFire !== libFire) return false;
  if (boqFire && libFire) {
    const boqMin = extractFireMinutes(boqText);
    const libMin = extractFireMinutes(libText);
    if (boqMin !== null && libMin !== null && boqMin !== libMin) return false;
  }
  return true;
}

// ─── Main Matching Engine V4.2 ────────────────────────────────────────────────

export function matchItem(
  item: BOQItem,
  library: RateLibraryItem[],
  _allItemsInFile: BOQItem[]
): MatchResult | null {

  const normItemNo = normalizeItemNo(item.item_no);
  const normUnit = normalizeUnit(item.unit);

  // ── Stage 1: item_no Exact ──────────────────────────────────────────────────
  if (normItemNo.length > 0) {
    for (const lib of library) {
      for (const alias of lib.item_name_aliases ?? []) {
        if (normalizeItemNo(alias) === normItemNo) {
          return { libraryItem: lib, confidence: 99, stage: 1, stageLabel: 'item_no_exact' };
        }
      }
    }
  }

  // ── Stage 2: Category + Unit Gate ──────────────────────────────────────────
  const unitFiltered = library.filter(lib => {
    const libUnit = normalizeUnit(lib.unit);
    return libUnit === normUnit || lib.unit === item.unit;
  });

  if (unitFiltered.length === 0) return null;

  const boqKeywords = generateKeywords(item.description);

  let best: { lib: RateLibraryItem; score: number } | null = null;

  // ── Stage 3: Description Match ─────────────────────────────────────────────
  for (const lib of unitFiltered) {
    if (!applyFireRatingGate(item.description, lib.standard_name_ar)) continue;

    const libKeywords = lib.keywords?.length
      ? lib.keywords
      : generateKeywords(lib.standard_name_ar);

    let score = jaccardSimilarity(boqKeywords, libKeywords) * 100;
    const ws = wordSimilarity(item.description, lib.standard_name_ar) * 100;
    score = Math.max(score, ws);

    score = applyThicknessGate(item.description, lib.standard_name_ar, score);

    if (score >= 85) {
      if (!best || score > best.score) {
        best = { lib, score };
      }
    }
  }

  if (best) {
    return {
      libraryItem: best.lib,
      confidence: Math.min(Math.round(best.score), 98),
      stage: 3,
      stageLabel: 'description_match',
    };
  }

  // ── Stage 4: Bundled Composite ────────────────────────────────────────────
  let compositeBest: { lib: RateLibraryItem; score: number } | null = null;

  for (const lib of unitFiltered) {
    if (!applyFireRatingGate(item.description, lib.standard_name_ar)) continue;

    const libKeywords = lib.keywords?.length
      ? lib.keywords
      : generateKeywords(lib.standard_name_ar);

    const itemNoSim = normItemNo.length > 0
      ? (normalizeItemNo(lib.standard_name_ar).includes(normItemNo) ? 0.5 : 0)
      : 0;
    const jaccardSim = jaccardSimilarity(boqKeywords, libKeywords);
    const wSim = wordSimilarity(item.description, lib.standard_name_ar);

    let score = (0.4 * itemNoSim + 0.3 * jaccardSim + 0.3 * wSim) * 100;
    score = applyThicknessGate(item.description, lib.standard_name_ar, score);

    if (score >= 75) {
      if (!compositeBest || score > compositeBest.score) {
        compositeBest = { lib, score };
      }
    }
  }

  if (compositeBest) {
    return {
      libraryItem: compositeBest.lib,
      confidence: Math.min(Math.round(compositeBest.score), 94),
      stage: 4,
      stageLabel: 'composite',
    };
  }

  return null;
}

export function runMatchingPipeline(
  items: BOQItem[],
  library: RateLibraryItem[]
): Map<string, MatchResult> {
  const results = new Map<string, MatchResult>();
  const approvedLibrary = library.filter(l => l.source_type === 'Approved');

  for (const item of items) {
    if (item.override_type === 'manual') continue;
    if (item.status === 'descriptive') continue;

    const match = matchItem(item, approvedLibrary, items);
    if (match) {
      results.set(item.id, match);
    }
  }

  return results;
}
