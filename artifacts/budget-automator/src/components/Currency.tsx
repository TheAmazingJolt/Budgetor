import { cn } from '@/lib/utils';

export function Currency({ value, className, showSign = false }: { value: number, className?: string, showSign?: boolean }) {
  const isPositive = value > 0;
  const isNegative = value < 0;
  
  const formatted = new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
  }).format(Math.abs(value));

  const prefix = showSign ? (isNegative ? '-' : isPositive ? '+' : '') : (isNegative ? '-' : '');

  return (
    <span className={cn(
      "font-mono font-medium tracking-tight",
      isPositive ? "text-[hsl(var(--success))]" : isNegative ? "text-[hsl(var(--danger))]" : "text-muted-foreground",
      className
    )}>
      {prefix}{formatted}
    </span>
  );
}
