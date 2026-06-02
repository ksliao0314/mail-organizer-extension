import * as React from 'react'
import { cn } from '@/lib/utils'

/* Card — Linear-inspired surface treatment.
 *
 * Changes from shadcn default:
 *   - radius `rounded-xl` (12px, was 0.625rem = 10px) — slightly more
 *     generous for outer cards, makes the corners feel intentional vs
 *     "stock"
 *   - shadow uses our warm-tinted var(--shadow-sm) via inline style
 *     instead of generic Tailwind shadow-sm (which is pure black at
 *     low alpha) — subtle but elevates the perceived quality
 *   - border softer at 50% to read as a hairline, not a hard edge
 *   - hover-elevation prepared (cards opt in by adding `hover:shadow-md`)
 *
 * The transition class is added but defaults to no animation; consumers
 * who want hover-lift add `hover:shadow-md hover:-translate-y-px`. */
export const Card = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, style, ...props }, ref) => (
  <div
    ref={ref}
    className={cn(
      'rounded-xl border border-border/70 bg-card text-card-foreground transition-all duration-200',
      className,
    )}
    style={{ boxShadow: 'var(--shadow-sm)', ...style }}
    {...props}
  />
))
Card.displayName = 'Card'

export const CardHeader = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, ...props }, ref) => (
  <div
    ref={ref}
    className={cn('flex flex-col space-y-1 p-4', className)}
    {...props}
  />
))
CardHeader.displayName = 'CardHeader'

/* CardTitle — was generic `text-sm font-semibold`. Now `text-base
 * font-semibold` with tighter tracking. h3-level by default (semantic
 * + the global h-tag tracking applies). */
export const CardTitle = React.forwardRef<
  HTMLHeadingElement,
  React.HTMLAttributes<HTMLHeadingElement>
>(({ className, ...props }, ref) => (
  <h3
    ref={ref}
    className={cn(
      'text-[15px] font-semibold leading-tight tracking-tight',
      className,
    )}
    {...props}
  />
))
CardTitle.displayName = 'CardTitle'

export const CardDescription = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, ...props }, ref) => (
  <div
    ref={ref}
    className={cn('text-xs text-muted-foreground', className)}
    {...props}
  />
))
CardDescription.displayName = 'CardDescription'

export const CardContent = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, ...props }, ref) => (
  <div ref={ref} className={cn('p-4 pt-0', className)} {...props} />
))
CardContent.displayName = 'CardContent'

export const CardFooter = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, ...props }, ref) => (
  <div
    ref={ref}
    className={cn('flex items-center p-4 pt-0', className)}
    {...props}
  />
))
CardFooter.displayName = 'CardFooter'
