import { Component, type ErrorInfo, type ReactNode } from 'react'
import { Button } from '@/components/ui/button'

// Root error boundary for the popup + options pages (2026-11-05).
//
// Without one, ANY render-time throw (a storage getter that throws in the
// OWA iframe, an unexpected SW payload shape) unmounts the whole React tree
// and leaves a blank white panel with no way out except closing it. The
// fallback says what happened in plain words and offers a reload — which
// re-runs the page and re-adopts any batch the service worker is still
// running, so nothing in flight is lost.

type Props = { children: ReactNode }
type State = { error: Error | null }

export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null }

  static getDerivedStateFromError(error: unknown): State {
    return { error: error instanceof Error ? error : new Error(String(error)) }
  }

  componentDidCatch(error: unknown, info: ErrorInfo) {
    console.error('[mail-organizer] UI crashed', error, info.componentStack)
  }

  render() {
    const { error } = this.state
    if (!error) return this.props.children
    return (
      <div
        role="alert"
        className="m-4 max-w-md rounded-md border border-amber-300 bg-amber-50 p-4 text-xs text-amber-950 space-y-2"
      >
        <div className="text-sm font-semibold">畫面發生錯誤</div>
        <p>
          Mail Organizer 的畫面遇到未預期的錯誤。重新載入通常就能恢復 —
          背景進行中的歸類批次不受影響，重新載入後會自動接回。
        </p>
        <details className="text-[10px] text-amber-900/80">
          <summary className="cursor-pointer select-none">技術細節</summary>
          <div className="mt-1 break-all font-mono">{error.message || error.name}</div>
        </details>
        <Button size="sm" onClick={() => window.location.reload()}>
          重新載入
        </Button>
      </div>
    )
  }
}
