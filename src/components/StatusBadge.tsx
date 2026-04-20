import type { BOQItemStatus } from '../types';

interface StatusBadgeProps {
  status: BOQItemStatus;
  size?: 'sm' | 'md';
}

const STATUS_CONFIG: Record<BOQItemStatus, { label: string; classes: string }> = {
  pending: {
    label: 'في الانتظار',
    classes: 'bg-amber-100 text-amber-800 border border-amber-200',
  },
  approved: {
    label: 'معتمد',
    classes: 'bg-emerald-100 text-emerald-800 border border-emerald-200',
  },
  manual: {
    label: 'يدوي',
    classes: 'bg-blue-100 text-blue-800 border border-blue-200',
  },
  stale_price: {
    label: 'سعر قديم',
    classes: 'bg-orange-100 text-orange-800 border border-orange-200',
  },
  descriptive: {
    label: 'وصفي',
    classes: 'bg-slate-100 text-slate-600 border border-slate-200',
  },
  needs_review: {
    label: 'يحتاج مراجعة',
    classes: 'bg-red-100 text-red-800 border border-red-200',
  },
};

export default function StatusBadge({ status, size = 'md' }: StatusBadgeProps) {
  const config = STATUS_CONFIG[status] ?? STATUS_CONFIG.pending;
  const sizeClasses = size === 'sm' ? 'px-1.5 py-0.5 text-xs' : 'px-2.5 py-1 text-xs';

  return (
    <span className={`inline-flex items-center rounded-full font-medium whitespace-nowrap ${sizeClasses} ${config.classes}`}>
      {config.label}
    </span>
  );
}
