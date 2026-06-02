import { describe, expect, it } from 'vitest'

// Pure-logic smoke test for the same-subject auto-propagate decision tree
// inside PlanScreen.handleItemChange. We replicate the relevant pieces of
// PlanScreen here (no React) to verify that:
//   1. A "complete" edit on a same-subject row triggers propagation.
//   2. Siblings with userTouched=true are skipped.
//   3. Subjects normalize via the Re:/Fw: stripper so user-perceived "same"
//      really matches.

type PlanAction = 'move' | 'delete' | 'new_folder' | 'skip'
type PlanItem = {
  emailId: string
  emailSubject: string
  action: PlanAction
  targetFolderPath?: string
  suggestedFolderName?: string
  suggestedParentPath?: string
  userTouched?: boolean
}

function normalizeSubject(s: string): string {
  let prev = ''
  let cur = s
  while (cur !== prev) {
    prev = cur
    cur = cur.replace(/^\s*(re|fw|fwd|轉寄|回覆|轉發|轉)[:：\s]+/i, '')
  }
  return cur.trim().replace(/\s+/g, ' ').toLowerCase()
}

function isComplete(item: PlanItem): boolean {
  return (
    item.action === 'delete' ||
    item.action === 'skip' ||
    (item.action === 'move' && !!item.targetFolderPath) ||
    (item.action === 'new_folder' &&
      !!item.suggestedFolderName?.trim() &&
      !!item.suggestedParentPath?.trim())
  )
}

function computeTargetIds(items: PlanItem[], source: PlanItem): string[] {
  if (!isComplete(source)) return []
  const norm = normalizeSubject(source.emailSubject)
  if (!norm) return []
  const sameSubject = items
    .filter((i) => i.emailId !== source.emailId && normalizeSubject(i.emailSubject) === norm)
    .map((i) => i.emailId)
  return sameSubject.filter((id) => !items.find((i) => i.emailId === id)?.userTouched)
}

describe('auto-propagate same subject — handleItemChange decision tree', () => {
  const baseItems: PlanItem[] = [
    { emailId: 'a', emailSubject: 'Re: 電子發票通知', action: 'move', targetFolderPath: '02/未分類' },
    { emailId: 'b', emailSubject: '電子發票通知', action: 'move', targetFolderPath: '02/未分類' },
    { emailId: 'c', emailSubject: 'Fwd: 電子發票通知', action: 'move', targetFolderPath: '02/未分類' },
    { emailId: 'd', emailSubject: '其他無關郵件', action: 'move', targetFolderPath: '02/未分類' },
  ]

  it('propagates to all same-normalized-subject siblings', () => {
    const source = { ...baseItems[0]!, action: 'delete' as const, userTouched: true }
    const targets = computeTargetIds(baseItems, source)
    expect(targets.sort()).toEqual(['b', 'c'])
    expect(targets).not.toContain('d')
    expect(targets).not.toContain('a')
  })

  it('strips Re:/Fwd:/轉寄: prefixes when computing the group', () => {
    const items = [
      { emailId: '1', emailSubject: '轉寄: 民事訴訟通知' },
      { emailId: '2', emailSubject: '民事訴訟通知' },
      { emailId: '3', emailSubject: '回覆: 民事訴訟通知' },
    ].map((x) => ({ ...x, action: 'move' as const, targetFolderPath: '/X' }))
    const source = { ...items[0]!, action: 'delete' as const, userTouched: true }
    const targets = computeTargetIds(items, source)
    expect(targets.sort()).toEqual(['2', '3'])
  })

  it('does NOT propagate when source is incomplete (move without target)', () => {
    const source = { emailId: 'a', emailSubject: 'Re: 電子發票通知', action: 'move' as const }
    const targets = computeTargetIds(baseItems, source)
    expect(targets).toEqual([])
  })

  it('does NOT propagate when source new_folder is missing parent path', () => {
    const source = {
      emailId: 'a',
      emailSubject: 'Re: 電子發票通知',
      action: 'new_folder' as const,
      suggestedFolderName: 'X',
      // parent missing
    }
    const targets = computeTargetIds(baseItems, source)
    expect(targets).toEqual([])
  })

  it('skips siblings marked userTouched=true', () => {
    const items = baseItems.map((i) => (i.emailId === 'b' ? { ...i, userTouched: true } : i))
    const source = { ...baseItems[0]!, action: 'delete' as const, userTouched: true }
    const targets = computeTargetIds(items, source)
    expect(targets).toEqual(['c'])
  })

  it('propagates when source has same target as siblings already (allows refinement of metadata)', () => {
    // Even when nothing visually changes for siblings (same target), the
    // decision tree still includes them — the actual bulkApply would no-op
    // visually but the userTouched filter is still applied at the source.
    const source = {
      ...baseItems[0]!,
      action: 'move' as const,
      targetFolderPath: '02/未分類', // same as siblings
      userTouched: true,
    }
    const targets = computeTargetIds(baseItems, source)
    expect(targets.sort()).toEqual(['b', 'c'])
  })

  it('handles single-item group (no siblings)', () => {
    const items = [{ emailId: 'lonely', emailSubject: 'unique subject', action: 'move' as const, targetFolderPath: '/X' }]
    const source = { ...items[0]!, action: 'delete' as const }
    expect(computeTargetIds(items, source)).toEqual([])
  })
})
