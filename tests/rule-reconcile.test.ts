import { describe, expect, it } from 'vitest'
import { diffSnapshots, reconcileRulesAgainstTree, snapshotOf } from '@/shared/rules'
import type { MailFolderNode, Rule } from '@/shared/types'

function rule(over: Partial<Rule> = {}): Rule {
  return {
    id: 'r-' + Math.random().toString(36).slice(2),
    type: 'domain',
    signal: 'example.com',
    targetFolderId: 'fid-abc',
    targetFolderPath: 'X/Y',
    confidence: 0.7,
    matchCount: 0,
    enabled: true,
    createdAt: '2026-01-01T00:00:00.000Z',
    source: 'user_manual',
    ...over,
  }
}

function leaf(id: string, displayName: string, parentPath = ''): MailFolderNode {
  return {
    id,
    displayName,
    parentFolderId: undefined,
    path: parentPath ? `${parentPath}/${displayName}` : displayName,
    children: [],
  }
}

describe('reconcileRulesAgainstTree', () => {
  it('leaves rules untouched when targetFolderId resolves to same path', () => {
    const tree: MailFolderNode[] = [leaf('fid-abc', 'Y', 'X')]
    // wrap inside X to match the path:
    const wrapped: MailFolderNode[] = [
      { id: 'fid-x', displayName: 'X', path: 'X', children: tree, parentFolderId: undefined },
    ]
    const r = rule({ targetFolderId: 'fid-abc', targetFolderPath: 'X/Y' })
    const result = reconcileRulesAgainstTree([r], wrapped)
    expect(result.summary.pathsUpdated).toBe(0)
    expect(result.summary.newlyOrphaned).toBe(0)
    expect(result.rules[0]).toEqual(r)
  })

  it('updates path when folder was renamed (Id still resolves)', () => {
    // Folder was Y, renamed to Y_new. Id stays the same.
    const renamed: MailFolderNode = {
      id: 'fid-abc',
      displayName: 'Y_new',
      path: 'X/Y_new',
      children: [],
      parentFolderId: 'fid-x',
    }
    const tree: MailFolderNode[] = [
      { id: 'fid-x', displayName: 'X', path: 'X', children: [renamed], parentFolderId: undefined },
    ]
    const r = rule({ targetFolderId: 'fid-abc', targetFolderPath: 'X/Y' })
    const result = reconcileRulesAgainstTree([r], tree)
    expect(result.summary.pathsUpdated).toBe(1)
    expect(result.summary.newlyOrphaned).toBe(0)
    expect(result.rules[0]!.targetFolderPath).toBe('X/Y_new')
    expect(result.rules[0]!.orphaned).toBeFalsy()
  })

  it('marks rule orphaned when target folder is deleted', () => {
    const tree: MailFolderNode[] = [leaf('fid-other', 'Other')]
    const r = rule({ targetFolderId: 'fid-gone', targetFolderPath: 'Deleted/Y' })
    const result = reconcileRulesAgainstTree([r], tree)
    expect(result.summary.newlyOrphaned).toBe(1)
    expect(result.rules[0]!.orphaned).toBe(true)
  })

  it('clears orphaned when folder reappears under same Id', () => {
    const tree: MailFolderNode[] = [leaf('fid-abc', 'Y', 'X')]
    const wrapped: MailFolderNode[] = [
      { id: 'fid-x', displayName: 'X', path: 'X', children: tree, parentFolderId: undefined },
    ]
    const r = rule({ targetFolderId: 'fid-abc', targetFolderPath: 'X/Y', orphaned: true })
    const result = reconcileRulesAgainstTree([r], wrapped)
    expect(result.summary.unorphaned).toBe(1)
    expect(result.rules[0]!.orphaned).toBe(false)
  })

  it('falls back to path lookup when targetFolderId is empty (legacy ai_confirmed rules)', () => {
    const tree: MailFolderNode[] = [leaf('fid-abc', 'Y', 'X')]
    const wrapped: MailFolderNode[] = [
      { id: 'fid-x', displayName: 'X', path: 'X', children: tree, parentFolderId: undefined },
    ]
    const r = rule({ targetFolderId: '', targetFolderPath: 'X/Y' })
    const result = reconcileRulesAgainstTree([r], wrapped)
    expect(result.rules[0]!.targetFolderId).toBe('fid-abc')
    expect(result.rules[0]!.orphaned).toBeFalsy()
  })

  it('does not double-count when rule is already orphaned and still missing', () => {
    const tree: MailFolderNode[] = [leaf('fid-other', 'Other')]
    const r = rule({ targetFolderId: 'fid-gone', targetFolderPath: 'Gone', orphaned: true })
    const result = reconcileRulesAgainstTree([r], tree)
    expect(result.summary.newlyOrphaned).toBe(0)
    expect(result.rules[0]).toEqual(r) // unchanged
  })

  it('handles mixed batch correctly', () => {
    const renamedY: MailFolderNode = {
      id: 'fid-y',
      displayName: 'Y_new',
      path: 'X/Y_new',
      children: [],
      parentFolderId: 'fid-x',
    }
    const keptZ: MailFolderNode = {
      id: 'fid-z',
      displayName: 'Z',
      path: 'X/Z',
      children: [],
      parentFolderId: 'fid-x',
    }
    const tree: MailFolderNode[] = [
      {
        id: 'fid-x',
        displayName: 'X',
        path: 'X',
        children: [renamedY, keptZ],
        parentFolderId: undefined,
      },
    ]
    const rules: Rule[] = [
      rule({ id: 'r1', targetFolderId: 'fid-y', targetFolderPath: 'X/Y' }), // renamed
      rule({ id: 'r2', targetFolderId: 'fid-z', targetFolderPath: 'X/Z' }), // unchanged
      rule({ id: 'r3', targetFolderId: 'fid-gone', targetFolderPath: 'X/Deleted' }), // orphan
    ]
    const result = reconcileRulesAgainstTree(rules, tree)
    expect(result.summary.pathsUpdated).toBe(1)
    expect(result.summary.newlyOrphaned).toBe(1)
    expect(result.rules.find((r) => r.id === 'r1')!.targetFolderPath).toBe('X/Y_new')
    expect(result.rules.find((r) => r.id === 'r2')!.targetFolderPath).toBe('X/Z')
    expect(result.rules.find((r) => r.id === 'r3')!.orphaned).toBe(true)
  })
})

// O1 (2026-11-06 r2). Round 1 never orphaned a target that "exists but isn't
// a filing target" — so a case folder deleted in OWA (now 刪除的郵件/甲, same
// id) kept routing mail every batch into an execute refusal. Now only
// Deleted Items ITSELF is a valid target outside the tree; everything else
// missing from a complete tree is orphaned, as in v1.0.8.
describe('reconcileRulesAgainstTree with well-known folders (O1)', () => {
  // The tree as built now: Junk in it; the trash, Drafts and Outbox not.
  const tree: MailFolderNode[] = [leaf('fid-junk', '垃圾郵件'), leaf('fid-case', '甲', '案件')]
  const trashFolder = { id: 'fid-trash', path: '刪除的郵件' }

  it('a USER rule on Deleted Items itself stays valid (auto-trash) — and one v1.0.8 orphaned is restored', () => {
    const live = rule({ targetFolderId: 'fid-trash', targetFolderPath: '刪除的郵件' })
    const old = rule({ targetFolderId: 'fid-trash', targetFolderPath: '刪除的郵件', orphaned: true })
    const out = reconcileRulesAgainstTree([live, old], tree, { trashFolder })
    expect(out.rules[0]).toBe(live)
    expect(out.rules[1]!.orphaned).toBe(false)
    expect(out.summary).toEqual({ pathsUpdated: 0, newlyOrphaned: 0, unorphaned: 1 })
  })

  // r3 (2026-11-06): every trash-root rule used to come back — so v1.0.7's
  // learned auto-trash rules (orphaned by shipped v1.0.8) resumed trashing.
  it('…only as the user’s intent: user_manual / ai_overridden stay valid, learned / scanned ones stay orphaned', () => {
    const at = (source: Rule['source'], orphaned?: boolean) =>
      rule({ id: source, targetFolderId: 'fid-trash', targetFolderPath: '刪除的郵件', source, ...(orphaned ? { orphaned } : {}) })
    const out = reconcileRulesAgainstTree(
      [at('user_manual', true), at('ai_overridden', true), at('ai_confirmed', true), at('auto_scan')],
      tree,
      { trashFolder },
    )
    expect(out.rules.map((r) => [r.id, r.orphaned ?? false])).toEqual([
      ['user_manual', false],
      ['ai_overridden', false],
      ['ai_confirmed', true],
      ['auto_scan', true],
    ])
    expect(out.summary).toMatchObject({ unorphaned: 2, newlyOrphaned: 1 })
  })

  it('…matched by path when the rule has no local id (synced from another mailbox)', () => {
    const r = rule({ targetFolderId: 'fid-other-mailbox-trash', targetFolderPath: '刪除的郵件' })
    const out = reconcileRulesAgainstTree([r], tree, { trashFolder })
    expect(out.rules[0]!.targetFolderId).toBe('fid-trash')
    expect(out.rules[0]!.orphaned ?? false).toBe(false)
  })

  it('a folder deleted in OWA (in the trash, same id) is orphaned, user_manual included', () => {
    const r = rule({ targetFolderId: 'fid-trashed-case', targetFolderPath: '案件/乙', source: 'user_manual' })
    const out = reconcileRulesAgainstTree([r], tree, { trashFolder })
    expect(out.rules[0]!.orphaned).toBe(true)
    expect(out.rules[0]!.targetFolderPath).toBe('案件/乙') // never re-pathed into the trash
  })

  it('Drafts / Outbox targets are orphaned like any missing folder', () => {
    const r = rule({ targetFolderId: 'fid-drafts', targetFolderPath: '草稿' })
    expect(reconcileRulesAgainstTree([r], tree, { trashFolder }).rules[0]!.orphaned).toBe(true)
  })

  it('an untrusted tree (partial lookup / legacy cache) leaves EVERY rule exactly as it is', () => {
    // A degraded fetch lists the trash subtree: 案件/乙 shows up as 刪除的郵件/乙.
    const degraded: MailFolderNode[] = [
      ...tree,
      { ...leaf('fid-trash', '刪除的郵件'), children: [leaf('fid-trashed-case', '乙', '刪除的郵件')] },
    ]
    const trashed = rule({ targetFolderId: 'fid-trashed-case', targetFolderPath: '案件/乙' })
    const gone = rule({ targetFolderId: 'fid-gone', targetFolderPath: '案件/丙' })
    const old = rule({ targetFolderId: 'fid-junk', targetFolderPath: '垃圾郵件', orphaned: true })
    const out = reconcileRulesAgainstTree([trashed, gone, old], degraded, { treeComplete: false })
    expect(out.rules).toEqual([trashed, gone, old])
    expect(out.rules[0]).toBe(trashed) // NOT rewritten to 刪除的郵件/乙
    expect(out.summary).toEqual({ pathsUpdated: 0, newlyOrphaned: 0, unorphaned: 0 })
  })

  it('Junk Email in the tree resolves like any folder (user_manual sender → 垃圾郵件)', () => {
    const r = rule({ type: 'sender', signal: 'spam@x.example', targetFolderId: 'fid-junk', targetFolderPath: '垃圾郵件', orphaned: true })
    const out = reconcileRulesAgainstTree([r], tree)
    expect(out.rules[0]!.orphaned).toBe(false)
  })
})
describe('snapshotOf / diffSnapshots — orphaned audit support', () => {
  it('includes orphaned in snapshot', () => {
    const r = rule({ orphaned: true })
    const snap = snapshotOf(r)
    expect(snap.orphaned).toBe(true)
  })

  it('diffSnapshots reports orphaned flip from undefined → true', () => {
    const r = rule({ orphaned: false })
    const before = snapshotOf({ ...r, orphaned: undefined })
    const after = snapshotOf({ ...r, orphaned: true })
    expect(diffSnapshots(before, after)).toContain('orphaned')
  })

  it('diffSnapshots treats undefined === false (no noise on first appearance)', () => {
    const r = rule()
    const a = snapshotOf({ ...r, orphaned: undefined })
    const b = snapshotOf({ ...r, orphaned: false })
    expect(diffSnapshots(a, b)).not.toContain('orphaned')
  })

  it('diffSnapshots picks up targetFolderPath change (reconcile rename)', () => {
    const before = snapshotOf(rule({ targetFolderPath: 'X/Y' }))
    const after = snapshotOf(rule({ targetFolderPath: 'X/Y_new' }))
    expect(diffSnapshots(before, after)).toContain('targetFolderPath')
  })

  it('no diff when only targetFolderId changes (snapshot intentionally drops it)', () => {
    // RuleSnapshot doesn't carry targetFolderId — id syncs without churning
    // the audit log are intentionally invisible (they happen on first scan
    // for legacy rules that were created with empty id).
    const before = snapshotOf(rule({ targetFolderId: '' }))
    const after = snapshotOf(rule({ targetFolderId: 'AAA' + 'x'.repeat(50) }))
    expect(diffSnapshots(before, after)).toEqual([])
  })
})
