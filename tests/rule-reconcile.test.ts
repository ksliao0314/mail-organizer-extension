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
