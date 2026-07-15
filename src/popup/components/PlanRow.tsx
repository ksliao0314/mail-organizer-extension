import { type ReactNode, useEffect, useMemo, useRef, useState } from 'react'
import { Button } from '@/components/ui/button'
import { Input } from '@/components/ui/input'
import { Badge } from '@/components/ui/badge'
import { FolderPicker } from './FolderPicker'
import { cn } from '@/lib/utils'
import { joinFolderPath } from '@/shared/outlook-api'
import type { MailFolderNode, PlanAction, PlanItem } from '@/shared/types'
import { useLayout } from '../layout-context'
import { AlertTriangle, Check, ChevronDown, ChevronRight, Sparkles, Tags } from 'lucide-react'

const ACTION_LABEL: Record<PlanAction, string> = {
  move: '移到',
  delete: '刪除',
  new_folder: '新建',
  skip: '保留',
}

const ACTION_TONE: Record<PlanAction, string> = {
  move: 'border-emerald-300 bg-emerald-50/70',
  delete: 'border-red-300 bg-red-50/70',
  new_folder: 'border-amber-300 bg-amber-50/70',
  skip: 'border-border bg-muted/40',
}

const ACTION_RAIL: Record<PlanAction, string> = {
  move: 'bg-emerald-500',
  delete: 'bg-red-500',
  new_folder: 'bg-amber-500',
  skip: 'bg-muted-foreground/40',
}

const ACTIONS: PlanAction[] = ['move', 'delete', 'new_folder', 'skip']

// Source badge — small chip identifying whether the routing decision came
// from a rule match, AI inference, or is unresolved. Exported so the
// wide-mode side-trace sidebar can render the same badge as the inline
// trace block without duplicating the JSX.
export function renderSourceBadge(source: PlanItem['source']): ReactNode {
  if (source === 'rule') {
    return (
      <Badge variant="outline" className="text-[9px]">
        <Tags className="size-2.5" />
        規則
      </Badge>
    )
  }
  if (source === 'ai') {
    return (
      <Badge variant="outline" className="text-[9px]">
        <Sparkles className="size-2.5" />
        AI
      </Badge>
    )
  }
  if (source === 'thread') {
    return (
      <Badge
        variant="outline"
        className="text-[9px] border-violet-300 bg-violet-50 text-violet-800"
      >
        對話延續
      </Badge>
    )
  }
  return (
    <Badge
      variant="outline"
      className="text-[9px] border-amber-300 bg-amber-50 text-amber-800"
    >
      待決
    </Badge>
  )
}

export type PlanRowProps = {
  item: PlanItem
  tree: MailFolderNode[]
  excludePrefixes: string[]
  selected: boolean
  /** Number of OTHER rows currently selected (excludes this one). */
  otherSelectedCount: number
  /** When true, draw a ring to indicate this row is the keyboard-nav target. */
  focused?: boolean
  /**
   * Increment to request the row scroll itself into view. Used by keyboard
   * navigation in PlanScreen so the focused row stays visible.
   */
  scrollIntoViewToken?: number
  /**
   * Increment to request the row toggle expanded/collapsed. Lets PlanScreen
   * drive Enter-key expansion without lifting expansion state out of the row.
   */
  toggleExpandToken?: number
  /**
   * Precomputed set of every valid folder path in the augmented tree.
   * Lifted from PlanScreen so we don't re-flatten the tree per-row per-render
   * (with 50 rows × 500 folders that was 25k ops on every keystroke).
   */
  validPaths: ReadonlySet<string>
  onChange: (next: PlanItem) => void
  onToggleSelect: () => void
  /** Copy this row's action+target onto all currently-selected rows + clear selection. */
  onPropagate: () => void
  /**
   * Disable/enable a rule the user wants to stop matching globally. Wired
   * from the trace block's inline button — only meaningful for rule-sourced
   * items where matchedRule is populated.
   */
  onToggleRule?: (ruleId: string, enabled: boolean) => Promise<void> | void
  /**
   * Called when the row should become the "focused" row (driving the wide-
   * mode side-trace sidebar). PlanScreen wires this to setFocusedIndex so
   * mouse hover updates the sidebar the same way keyboard j/k does. Only
   * passed in wide mode — narrow mode falls back to keyboard-only.
   */
  onActivate?: () => void
}

export function PlanRow({
  item,
  tree,
  excludePrefixes,
  selected,
  otherSelectedCount,
  focused,
  scrollIntoViewToken,
  toggleExpandToken,
  validPaths,
  onChange,
  onToggleSelect,
  onPropagate,
  onToggleRule,
  onActivate,
}: PlanRowProps) {
  const [expanded, setExpanded] = useState(false)
  const rowRef = useRef<HTMLLIElement | null>(null)
  const { isWide } = useLayout()

  // Imperatively trigger expand toggle when PlanScreen bumps the token.
  // The token prop only exists while this row is focused, so the FIRST
  // value a row instance sees just means "focus arrived" — sync it
  // silently. Only a CHANGE between two defined values is a real Enter
  // press. (Audit P1: comparing against an undefined ref made focus gain
  // itself count as Enter — j/k auto-expanded every row it landed on, and
  // the first click on a newly-focused row was swallowed by the bubbled
  // toggle.)
  const lastExpandToken = useRef<number | undefined>(undefined)
  useEffect(() => {
    const prev = lastExpandToken.current
    lastExpandToken.current = toggleExpandToken
    if (toggleExpandToken === undefined || prev === undefined) return
    if (toggleExpandToken === prev) return
    setExpanded((v) => !v)
  }, [toggleExpandToken])

  // Scroll into view when the keyboard-nav cursor lands on this row.
  // `nearest` keeps the page mostly still — only nudges when row is offscreen.
  const lastScrollToken = useRef<number | undefined>(undefined)
  useEffect(() => {
    if (scrollIntoViewToken === undefined) return
    if (scrollIntoViewToken === lastScrollToken.current) return
    lastScrollToken.current = scrollIntoViewToken
    rowRef.current?.scrollIntoView({ block: 'nearest', behavior: 'smooth' })
  }, [scrollIntoViewToken])

  // Source row must be in a "complete" state before its target can sensibly
  // be propagated — otherwise we'd copy emptiness to other rows.
  const canPropagate =
    item.action === 'delete' ||
    item.action === 'skip' ||
    (item.action === 'move' && !!item.targetFolderPath) ||
    (item.action === 'new_folder' &&
      !!item.suggestedFolderName?.trim() &&
      !!item.suggestedParentPath?.trim())

  // Pre-execute path validation: warn if move target or new_folder parent
  // doesn't resolve against the current tree (would error at execute time).
  //
  // "pending:" sentinels point at a folder that a *sibling* new_folder item
  // will create earlier in the same batch. We do NOT blanket-skip the check
  // for them: `validPaths` is built from buildAugmentedTree, which already
  // injects a virtual node for every live new_folder destination. So while
  // the backing new_folder still exists the pending path IS in validPaths and
  // the normal pathExists check passes. But if the user later changes that
  // source row's action away from new_folder (or edits its name/parent), the
  // virtual node disappears — the pending path is no longer in validPaths and
  // we now (correctly) warn, instead of letting the dependent move error red
  // at execute time with no prior heads-up.
  const targetIssue = useMemo<string | null>(() => {
    const pathExists = (p: string | undefined) => !!p && validPaths.has(p)
    if (item.action === 'move' && item.targetFolderPath) {
      if (!pathExists(item.targetFolderPath)) return '目標資料夾路徑不在當前資料夾樹中、執行時會失敗'
    }
    if (item.action === 'new_folder' && item.suggestedParentPath) {
      if (!pathExists(item.suggestedParentPath)) return '父資料夾路徑不在當前資料夾樹中、無法在那建立'
    }
    return null
  }, [item.action, item.targetFolderPath, item.suggestedParentPath, validPaths])

  function setAction(action: PlanAction) {
    if (action === item.action) return
    const next: PlanItem = { ...item, action, userTouched: true }
    // Reset action-specific fields
    if (action !== 'move') {
      next.targetFolderId = undefined
      next.targetFolderPath = undefined
    }
    if (action !== 'new_folder') {
      next.suggestedFolderName = undefined
      next.suggestedParentPath = undefined
    }
    // Once the user touches a rule-sourced item, the original rule no longer
    // explains the outcome. Detach source + ruleId so post-execute feedback
    // (bumpRuleHit) doesn't credit a rule that wasn't actually applied.
    next.source = 'ai'
    next.ruleId = undefined
    onChange(next)
  }

  function setTarget(node: MailFolderNode) {
    onChange({
      ...item,
      targetFolderId: node.id,
      targetFolderPath: node.path,
      userTouched: true,
      // Editing the target invalidates a rule-sourced attribution too.
      source: item.source === 'rule' ? 'ai' : item.source,
      ruleId: item.source === 'rule' ? undefined : item.ruleId,
    })
  }

  function setNewFolderName(name: string) {
    onChange({ ...item, suggestedFolderName: name, userTouched: true })
  }

  function setNewFolderParent(node: MailFolderNode) {
    onChange({ ...item, suggestedParentPath: node.path, userTouched: true })
  }

  const display =
    item.action === 'move'
      ? item.targetFolderPath
      : item.action === 'new_folder'
      ? `${joinFolderPath(item.suggestedParentPath, item.suggestedFolderName ?? '?')}（新建）`
      : item.action === 'delete'
      ? '永久刪除'
      : '保留在收件夾'

  // Only unresolved items earn an inline badge in the collapsed view —
  // rule/AI tone is carried by the action rail; full source detail is
  // shown on expand. Keeps rows scannable in 50-100 batches.
  const collapsedAttentionBadge =
    item.source === 'unresolved' ? (
      <Badge
        variant="outline"
        className="text-[9px] border-amber-300 bg-amber-50 text-amber-800"
        title="AI 信心低於門檻、未產生明確分類，請你逐封決定"
      >
        待決
      </Badge>
    ) : null

  const sourceDetailBadge = renderSourceBadge(item.source)

  // Show the confidence raw number only inside the trace block (on expand).
  // In the collapsed row we surface only the truly-low case as a chip — a
  // running list of 0.62 / 0.78 / 0.83 numbers is visual noise and the
  // user can dig into the exact value when they care.
  const lowConfidence = item.source !== 'rule' && item.confidence < 0.5

  const fromCompact = (() => {
    const raw = item.emailFrom?.trim() ?? ''
    if (!raw) return '(寄件人未知)'
    const at = raw.lastIndexOf('@')
    return at >= 0 && at < raw.length - 1 ? raw.slice(at + 1) : raw
  })()

  return (
    <li
      ref={rowRef}
      onClick={onActivate}
      className={cn(
        'relative rounded-md border overflow-hidden transition-colors',
        ACTION_TONE[item.action],
        selected && 'ring-2 ring-foreground ring-offset-1 ring-offset-background',
        focused && !selected && 'ring-2 ring-foreground/50 ring-offset-1 ring-offset-background',
      )}
    >
      <div className={cn('absolute left-0 top-0 bottom-0 w-1', ACTION_RAIL[item.action])} />
      <div className="flex items-start gap-2 pl-3 pr-3 py-2">
        {/* Checkbox — own click area, doesn't toggle the expand */}
        <div onClick={(e) => e.stopPropagation()} className="pt-0.5">
          <input
            type="checkbox"
            checked={selected}
            onChange={onToggleSelect}
            className="accent-foreground cursor-pointer"
            aria-label="選取此項以套用批次動作"
          />
        </div>
        <button
          type="button"
          onClick={() => setExpanded((v) => !v)}
          className="flex-1 min-w-0 flex items-start gap-2 text-left -mx-1 px-1 py-0.5 -my-0.5 rounded hover:bg-black/[0.02]"
        >
          <div className="pt-0.5">
            {expanded ? <ChevronDown className="size-3.5 text-muted-foreground" /> : <ChevronRight className="size-3.5 text-muted-foreground" />}
          </div>
          <div className="flex-1 min-w-0 space-y-0.5">
            {isWide ? (
              // Wide (reading-pane mode): subject + from share one line with
              // plenty of horizontal room, action + path on the second line
              // — total two lines instead of the cramped three-line stack
              // used in the 540px Chrome popup.
              <>
                <div className="flex items-baseline gap-3">
                  <div className="font-medium text-sm flex-1 min-w-0 truncate">
                    {item.emailSubject || '（無主旨）'}
                  </div>
                  <div
                    className="font-mono text-[11px] text-muted-foreground shrink-0 max-w-[280px] truncate"
                    title={item.emailFrom}
                  >
                    {item.emailFrom}
                  </div>
                  {lowConfidence && (
                    <Badge
                      variant="outline"
                      className="text-[9px] border-amber-300 bg-amber-50 text-amber-800 gap-0.5 shrink-0"
                      title={`AI 信心 ${item.confidence.toFixed(2)} — 展開看判斷依據`}
                    >
                      <AlertTriangle className="size-2.5" />
                      低信心
                    </Badge>
                  )}
                </div>
                <div className="flex items-baseline gap-2 text-[12px]">
                  <span className="font-medium shrink-0">{ACTION_LABEL[item.action]}</span>
                  <span
                    className={cn('flex-1 min-w-0 truncate', targetIssue ? 'text-red-700' : 'text-muted-foreground')}
                  >
                    {display ?? '—'}
                  </span>
                  {targetIssue && (
                    <AlertTriangle className="size-3.5 text-red-600 shrink-0" aria-label={targetIssue} />
                  )}
                  {collapsedAttentionBadge && (
                    <span className="shrink-0">{collapsedAttentionBadge}</span>
                  )}
                </div>
                {targetIssue && (
                  <div className="flex items-center gap-1 text-[10px] text-red-700 leading-tight">
                    <AlertTriangle className="size-3 shrink-0" />
                    <span>{targetIssue}</span>
                  </div>
                )}
              </>
            ) : (
              // Narrow (Chrome toolbar popup at 540px): three-line stack.
              <>
                <div className="flex items-baseline justify-between gap-2">
                  <div className="font-medium text-xs truncate">{item.emailSubject || '（無主旨）'}</div>
                  {lowConfidence && (
                    <Badge
                      variant="outline"
                      className="text-[9px] border-amber-300 bg-amber-50 text-amber-800 gap-0.5 shrink-0"
                      title={`AI 信心 ${item.confidence.toFixed(2)} — 展開看判斷依據`}
                    >
                      <AlertTriangle className="size-2.5" />
                      低信心
                    </Badge>
                  )}
                </div>
                <div
                  className="font-mono text-[10px] text-muted-foreground truncate"
                  title={item.emailFrom}
                >
                  {fromCompact}
                </div>
                <div className="flex items-baseline gap-1.5 text-[11px]">
                  <span className="font-medium">{ACTION_LABEL[item.action]}</span>
                  <span className={cn('truncate', targetIssue ? 'text-red-700' : 'text-muted-foreground')}>
                    {display ?? '—'}
                  </span>
                  {targetIssue && (
                    <AlertTriangle className="size-3 text-red-600 shrink-0" aria-label={targetIssue} />
                  )}
                  {collapsedAttentionBadge && (
                    <span className="ml-auto shrink-0">{collapsedAttentionBadge}</span>
                  )}
                </div>
                {targetIssue && (
                  <div className="flex items-center gap-1 text-[10px] text-red-700 leading-tight">
                    <AlertTriangle className="size-3 shrink-0" />
                    <span>{targetIssue}</span>
                  </div>
                )}
              </>
            )}
          </div>
        </button>
      </div>

      {expanded && (
        <div className="px-3 pb-3 pl-6 space-y-2 border-t border-black/5 pt-2">
          {item.bodyPreview && item.bodyPreview.trim().length > 0 && (
            <div
              className="rounded-md border border-border bg-muted/40 px-2.5 py-2 text-[11px] leading-relaxed text-muted-foreground whitespace-pre-wrap break-words"
              aria-label="信件摘要"
            >
              {item.bodyPreview.slice(0, 350)}
              {item.bodyPreview.length > 350 && '…'}
            </div>
          )}
          <RuleTraceBlock
            item={item}
            sourceDetailBadge={sourceDetailBadge}
            onToggleRule={onToggleRule}
          />
          <div className="flex gap-1 flex-wrap">
            {ACTIONS.map((a) => (
              <button
                key={a}
                type="button"
                onClick={() => setAction(a)}
                className={cn(
                  'rounded-md border px-2 py-1 text-[10px] font-medium transition-colors',
                  a === item.action
                    ? 'border-foreground bg-foreground text-background'
                    : 'border-border bg-background hover:bg-accent',
                )}
              >
                {ACTION_LABEL[a]}
              </button>
            ))}
          </div>

          {item.action === 'move' && (
            <div className="space-y-1">
              <label className="text-[10px] text-muted-foreground">目標資料夾</label>
              <FolderPicker
                tree={tree}
                value={item.targetFolderPath}
                excludePrefixes={excludePrefixes}
                onSelect={setTarget}
              />
            </div>
          )}

          {item.action === 'new_folder' && (
            <div className="space-y-2">
              <div className="space-y-1">
                <label className="text-[10px] text-muted-foreground">新資料夾名稱</label>
                <Input
                  value={item.suggestedFolderName ?? ''}
                  onChange={(e) => setNewFolderName(e.target.value)}
                  placeholder="例：XX公司勞資爭議案"
                  className="text-xs"
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] text-muted-foreground">父資料夾位置</label>
                <FolderPicker
                  tree={tree}
                  value={item.suggestedParentPath}
                  excludePrefixes={excludePrefixes}
                  onSelect={setNewFolderParent}
                />
              </div>
            </div>
          )}

          {item.reason && (
            <div className="text-[10px] text-muted-foreground italic">{item.reason}</div>
          )}

          <div className="flex items-center justify-between gap-2 pt-1">
            <div className="flex-1">
              {otherSelectedCount > 0 && (
                <Button
                  size="sm"
                  variant="default"
                  disabled={!canPropagate}
                  onClick={onPropagate}
                  title={
                    canPropagate
                      ? `把這張卡的目標套用到其他 ${otherSelectedCount} 張勾選的卡，套用後自動清空勾選`
                      : '這張卡的目標尚未指定、無法套用'
                  }
                >
                  全部套用到勾選 {otherSelectedCount} 項
                </Button>
              )}
            </div>
            <Button size="sm" variant="ghost" onClick={() => setExpanded(false)}>收起</Button>
          </div>
        </div>
      )}
    </li>
  )
}


// ---- Rule trace block ----------------------------------------------------
//
// Renders the "why" panel inside an expanded PlanRow. For rule-sourced
// items: shows the matched rule's signal/target/source/usage stats plus an
// inline "disable rule" button so the user can kill noisy rules without
// leaving the plan. For AI items: shows confidence + reason. For
// unresolved: tells the user the AI couldn't pick.

const RULE_SOURCE_LABEL: Record<string, string> = {
  user_manual: '手動建立',
  auto_scan: '初始掃描',
  ai_confirmed: 'AI 確認',
  ai_overridden: 'AI 覆蓋學習 (你之前改過)',
}

const RULE_TYPE_LABEL: Record<string, string> = {
  domain: '網域',
  sender: '寄件人',
  case_code: '案件代號',
  subject_keyword: '主旨關鍵字',
  compound: '複合條件',
}

function formatRelativeTime(iso: string | undefined): string | null {
  if (!iso) return null
  const ms = Date.parse(iso)
  if (!Number.isFinite(ms)) return null
  const diff = Date.now() - ms
  if (diff < 0) return '剛剛'
  const sec = Math.floor(diff / 1000)
  if (sec < 60) return `${sec} 秒前`
  const min = Math.floor(sec / 60)
  if (min < 60) return `${min} 分前`
  const hr = Math.floor(min / 60)
  if (hr < 24) return `${hr} 小時前`
  const day = Math.floor(hr / 24)
  if (day < 30) return `${day} 天前`
  const mon = Math.floor(day / 30)
  if (mon < 12) return `${mon} 月前`
  return `${Math.floor(mon / 12)} 年前`
}

export function RuleTraceBlock({
  item,
  sourceDetailBadge,
  onToggleRule,
}: {
  item: PlanItem
  sourceDetailBadge: ReactNode
  onToggleRule?: (ruleId: string, enabled: boolean) => Promise<void> | void
}) {
  const [localDisabled, setLocalDisabled] = useState(false)
  const [toggleFailed, setToggleFailed] = useState(false)
  const mr = item.matchedRule

  // Compute AI-override hint: when user changed away from what AI suggested.
  const aiOriginalTarget = item.aiOriginalAction === 'move'
    ? item.aiOriginalTargetFolderPath
    : item.aiOriginalAction === 'new_folder' && item.aiOriginalSuggestedParentPath && item.aiOriginalSuggestedFolderName
      ? joinFolderPath(item.aiOriginalSuggestedParentPath, item.aiOriginalSuggestedFolderName)
      : undefined
  const finalTarget = item.action === 'move'
    ? item.targetFolderPath
    : item.action === 'new_folder' && item.suggestedParentPath && item.suggestedFolderName
      ? joinFolderPath(item.suggestedParentPath, item.suggestedFolderName)
      : undefined
  const userOverrodeAi =
    item.aiOriginalAction !== undefined &&
    (item.aiOriginalAction !== item.action || aiOriginalTarget !== finalTarget)

  return (
    <div className="rounded border border-border/60 bg-muted/30 px-2.5 py-2 space-y-1.5 text-[10px]">
      <div className="flex items-center gap-2 text-muted-foreground">
        <span className="font-medium text-foreground">判斷依據</span>
        {sourceDetailBadge}
        <span className="font-mono tabular-nums">信心 {item.confidence.toFixed(2)}</span>
      </div>

      {item.source === 'rule' && mr && (
        <div className="space-y-1">
          <div className="font-mono text-muted-foreground">
            {RULE_TYPE_LABEL[mr.type] ?? mr.type}: <span className="text-foreground">{mr.signal}</span>
          </div>
          <div className="text-muted-foreground">
            來源: {RULE_SOURCE_LABEL[mr.source] ?? mr.source}
            {' · '}
            命中 <span className="font-mono tabular-nums">{mr.matchCount}</span> 次
            {formatRelativeTime(mr.lastUsedAt) && ` · 上次 ${formatRelativeTime(mr.lastUsedAt)}`}
          </div>
          {item.reason && (
            <div className="italic text-muted-foreground">{item.reason}</div>
          )}
          {onToggleRule && mr.enabled && !localDisabled && (
            <div className="pt-1">
              <Button
                size="sm"
                variant="outline"
                className="h-6 px-2 text-[10px]"
                onClick={async () => {
                  setLocalDisabled(true)
                  setToggleFailed(false)
                  try {
                    await onToggleRule(mr.id, false)
                  } catch {
                    // Roll back optimistic flag so the button reappears, and
                    // surface a visible hint so the user knows the toggle
                    // didn't actually take effect (rather than silently
                    // leaving them thinking the rule is disabled).
                    setLocalDisabled(false)
                    setToggleFailed(true)
                  }
                }}
                title="停用此規則,以後不再自動命中"
              >
                停用此規則
              </Button>
              {toggleFailed && (
                <span className="ml-2 text-red-700 italic">停用失敗、請再試一次</span>
              )}
            </div>
          )}
          {localDisabled && (
            <div className="flex items-center gap-1 text-emerald-700 italic">
              <Check className="size-3 shrink-0" />
              <span>已停用、未來不再自動命中</span>
            </div>
          )}
        </div>
      )}

      {item.source === 'ai' && (
        <div className="space-y-1">
          {item.reason && (
            <div className="italic text-muted-foreground">{item.reason}</div>
          )}
          {userOverrodeAi && aiOriginalTarget && (
            <div className="text-muted-foreground">
              AI 原本建議 <span className="font-mono">{aiOriginalTarget}</span>、你改成現在的目標
            </div>
          )}
          <div className="text-[9px] text-muted-foreground">
            執行後,AI 會根據你的選擇自動學習產生規則
          </div>
        </div>
      )}

      {item.source === 'thread' && (
        <div className="space-y-1">
          <div className="text-muted-foreground">
            {item.threadMatch?.kind === 'convId'
              ? '同一對話 thread 的歷史信件曾被歸到:'
              : '相同主旨的歷史信件曾被歸到:'}{' '}
            <span className="font-mono text-foreground">
              {item.threadMatch?.previousFolderPath ?? '(已忘記)'}
            </span>
          </div>
          {item.threadMatch?.kind === 'subject' && (
            <div className="text-[9px] text-muted-foreground">
              主旨匹配為較寬鬆的 fallback;若該主旨後續被歸到不同資料夾(衝突),則此機制自動停止建議。
            </div>
          )}
          {item.reason && (
            <div className="italic text-muted-foreground">{item.reason}</div>
          )}
        </div>
      )}

      {item.source === 'unresolved' && (
        <div className="space-y-1">
          {item.reason && (
            <div className="italic text-muted-foreground">{item.reason}</div>
          )}
          <div className="text-muted-foreground">
            AI 信心過低,沒給出明確分類,請你逐封決定
          </div>
        </div>
      )}
    </div>
  )
}
