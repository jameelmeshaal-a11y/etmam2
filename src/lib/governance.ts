import type { RateLibraryItem } from '../types';

export interface GovernanceViolation {
  wall: number;
  message: string;
  details?: Record<string, unknown>;
}

export class GovernanceError extends Error {
  wall: number;
  details?: Record<string, unknown>;
  constructor(violation: GovernanceViolation) {
    super(violation.message);
    this.name = 'GovernanceError';
    this.wall = violation.wall;
    this.details = violation.details;
  }
}

export function enforceWall4(rate: number, libraryItem: RateLibraryItem): void {
  if (rate < libraryItem.rate_min || rate > libraryItem.rate_max) {
    throw new GovernanceError({
      wall: 4,
      message: `Rate ${rate} is outside governance bounds [${libraryItem.rate_min}, ${libraryItem.rate_max}] for "${libraryItem.standard_name_ar}"`,
      details: {
        rate,
        rate_min: libraryItem.rate_min,
        rate_max: libraryItem.rate_max,
        item: libraryItem.standard_name_ar,
      },
    });
  }
}

export function enforceWall5(workbookTotal: number, dbTotal: number): void {
  if (dbTotal === 0) return;
  const variance = Math.abs((workbookTotal - dbTotal) / dbTotal) * 100;
  if (variance > 2.5) {
    throw new GovernanceError({
      wall: 5,
      message: `Export variance ${variance.toFixed(2)}% exceeds the ±2.5% hard limit. Export blocked.`,
      details: {
        workbookTotal,
        dbTotal,
        variance,
      },
    });
  }
}

export function enforceManualLock(overrideType: string | null, itemNo: string): void {
  if (overrideType === 'manual') {
    throw new GovernanceError({
      wall: 2,
      message: `Item ${itemNo} is permanently locked by manual override and cannot be auto-repriced.`,
      details: { itemNo },
    });
  }
}

export function checkRateBounds(rate: number, rateMin: number, rateMax: number): boolean {
  return rate >= rateMin && rate <= rateMax;
}
