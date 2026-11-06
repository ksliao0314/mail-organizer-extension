import { useCallback, useEffect, useState, useSyncExternalStore, type ReactNode } from 'react'
import { Loader2 } from 'lucide-react'
import { getParentBridge, type ParentBridge } from '@/shared/parent-bridge'

// The popup / options pages, when FRAMED, render their UI only once the FAB
// content script has paired (2026-11-06).
//
// They are web-accessible to the Outlook origins (owa-fab shows the popup in
// its panel and the options page in its settings modal, so neither can leave
// web_accessible_resources). That lets an Outlook page script embed them
// itself — as fully privileged extension pages — and overlay or disguise
// them to trick clicks (clickjacking): 執行 a plan, 全部刪除 rules, turn off
// sync. Pairing (shared/parent-bridge.ts) only protects the message channel.
// A page script can't complete pairing — the nonce comes from chrome.runtime
// — so "paired" is what tells our own panel from a foreign frame: until then
// a framed page shows a neutral placeholder, never the UI.
//
// The same wait catches a content script orphaned by an extension update
// (2026-11-06): it can't pair either (no runtime for a nonce; a pre-nonce
// version sends hellos the bridge refuses), so the panel's close / Esc /
// 近日活動 navigation would silently do nothing. Once the content script's
// re-offer schedule is over, the placeholder turns into a notice saying so.
// The toolbar popup and a standalone tab are top-level: no gate.

/**
 * When to stop waiting: past the FAB's re-offer schedule (owa-fab.ts
 * HELLO_RETRY_DELAYS_MS ≈ 7.5 s, each offer after an SW round trip). A
 * pairing that still completes later shows the UI anyway.
 */
export const PAIRING_NOTICE_AFTER_MS = 10_000
/** A normal pairing takes well under this — don't flash a placeholder text. */
const PLACEHOLDER_TEXT_AFTER_MS = 400

type Props = {
  children: ReactNode
  /** Injectable for tests; defaults to this page's bridge. */
  bridge?: ParentBridge | null
  noticeAfterMs?: number
}

export function BridgeGate({ children, bridge, noticeAfterMs = PAIRING_NOTICE_AFTER_MS }: Props) {
  const b = bridge === undefined ? getParentBridge() : bridge
  const gated = !!b && b.isEmbedded()
  const subscribe = useCallback((onChange: () => void) => (b ? b.onConnected(onChange) : () => {}), [b])
  const connected = useSyncExternalStore(subscribe, () => !!b && b.isConnected())
  const waiting = gated && !connected

  const [stage, setStage] = useState<'blank' | 'placeholder' | 'notice'>('blank')
  useEffect(() => {
    if (!waiting) return
    const text = window.setTimeout(() => setStage((s) => (s === 'blank' ? 'placeholder' : s)), PLACEHOLDER_TEXT_AFTER_MS)
    const notice = window.setTimeout(() => setStage('notice'), noticeAfterMs)
    return () => {
      window.clearTimeout(text)
      window.clearTimeout(notice)
    }
  }, [waiting, noticeAfterMs])

  if (!waiting) return <>{children}</>
  if (stage === 'notice') return <StaleConnectionNotice />
  return (
    <div
      role="status"
      aria-live="polite"
      aria-busy="true"
      className="flex min-h-[200px] items-center justify-center gap-2 p-6 text-xs text-muted-foreground"
    >
      {stage === 'placeholder' && (
        <>
          <Loader2 className="size-3.5 animate-spin" aria-hidden="true" />
          正在連線 Outlook 分頁…
        </>
      )}
    </div>
  )
}

function StaleConnectionNotice() {
  return (
    <div
      role="alert"
      className="m-4 max-w-md rounded-md border border-amber-300 bg-amber-50 p-4 text-xs text-amber-950 space-y-2"
    >
      <div className="text-sm font-semibold">
        與 Outlook 分頁的連線已失效（擴充功能已更新），請重新整理 Outlook 分頁
      </div>
      <p>重新整理（F5 / ⌘R）後再打開這個面板即可。現在可以先點面板外的區域把它關閉。</p>
    </div>
  )
}
