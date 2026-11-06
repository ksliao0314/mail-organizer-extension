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
//
// The reload also restores the saved plan (2026-11-06), so a crash caused by
// THAT plan's data came straight back on every reload. The popup therefore
// passes `onDiscardState`, a second way out that drops the saved plan first
// (and a finished batch's result screen — popup/crash-recovery.ts).

/** Give up waiting for the discard and reload anyway — a hung SW must not trap the user. */
const DISCARD_TIMEOUT_MS = 3000

type Props = {
  children: ReactNode
  /**
   * Popup only: throw away the state a reload would restore (the saved plan
   * and the classify run feeding it). Shown as 「捨棄暫存計畫並重新載入」.
   */
  onDiscardState?: () => Promise<unknown>
  /** Injectable for tests; defaults to `window.location.reload()`. */
  reload?: () => void
}
type State = { error: Error | null; discarding: boolean }

export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null, discarding: false }

  static getDerivedStateFromError(error: unknown): Partial<State> {
    return { error: error instanceof Error ? error : new Error(String(error)) }
  }

  componentDidCatch(error: unknown, info: ErrorInfo) {
    console.error('[mail-organizer] UI crashed', error, info.componentStack)
  }

  private reload = () => {
    if (this.props.reload) this.props.reload()
    else window.location.reload()
  }

  private discardAndReload = async () => {
    const discard = this.props.onDiscardState
    if (!discard || this.state.discarding) return
    this.setState({ discarding: true })
    let timer: ReturnType<typeof setTimeout> | undefined
    try {
      await Promise.race([
        discard(),
        new Promise((resolve) => {
          timer = setTimeout(resolve, DISCARD_TIMEOUT_MS)
        }),
      ])
    } catch (e) {
      // Reload anyway — staying on the error screen helps nobody.
      console.warn('[mail-organizer] discarding saved state failed', e)
    } finally {
      if (timer !== undefined) clearTimeout(timer)
    }
    this.reload()
  }

  render() {
    const { error, discarding } = this.state
    if (!error) return this.props.children
    const canDiscard = !!this.props.onDiscardState
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
        {canDiscard && (
          <p>
            如果重新載入後還是同樣的錯誤，可以捨棄暫存的歸類計畫再重新載入：
            這一批尚未執行的調整、進行中的 AI 分類，以及已完成批次的結果畫面（含撤回）會被捨棄；
            未執行的郵件仍在收件匣，可以重新歸類。執行中的批次不受影響。
          </p>
        )}
        <details className="text-[10px] text-amber-900/80">
          <summary className="cursor-pointer select-none">技術細節</summary>
          <div className="mt-1 break-all font-mono">{error.message || error.name}</div>
        </details>
        <div className="flex flex-wrap gap-2">
          <Button size="sm" onClick={this.reload} disabled={discarding}>
            重新載入
          </Button>
          {canDiscard && (
            <Button
              size="sm"
              variant="outline"
              onClick={() => void this.discardAndReload()}
              disabled={discarding}
            >
              {discarding ? '正在捨棄…' : '捨棄暫存計畫並重新載入'}
            </Button>
          )}
        </div>
      </div>
    )
  }
}
