import { describe, expect, it } from 'vitest'
import { filterFolderActivity } from '@/shared/folder-activity-filter'
import type { FolderActivity } from '@/shared/types'

function row(folderPath: string): FolderActivity {
  return {
    folderId: 'id-' + folderPath,
    folderPath,
    lastActiveAt: '2026-05-22T00:00:00Z',
    recentCount: 1,
  }
}

const sample: FolderActivity[] = [
  row('Active/Foo'),
  row('Active/Bar'),
  row('Archive/Old'),
  row('Inbox/SubA/重要客戶'),
  row('Inbox/SubB/重要客戶'),
]

describe('filterFolderActivity', () => {
  // ---- Regression bug fixed 2026-05-22 ----------------------------------
  //
  // Fresh installs default to `recentActivityIncludePrefixes: []` and
  // `recentActivityIncludeLeafNames: []`. Before the fix, the filter
  // treated empty allowlists as "deny all" → IdleScreen 近日活動 stayed
  // empty after the very first batch → user stuck on the start banner
  // thinking nothing happened. Fix: empty allowlists ⇒ no filtering.
  describe('empty-allowlist fallback', () => {
    it('returns all rows when both prefixes and leafNames are empty', () => {
      const out = filterFolderActivity(sample, [], new Set())
      expect(out).toHaveLength(sample.length)
      expect(out.map((r) => r.folderPath)).toEqual(sample.map((r) => r.folderPath))
    })

    it('returns a copy, not the same array reference (mutation safety)', () => {
      const out = filterFolderActivity(sample, [], new Set())
      expect(out).not.toBe(sample)
    })

    it('handles empty input + empty filters → empty result', () => {
      expect(filterFolderActivity([], [], new Set())).toEqual([])
    })
  })

  describe('prefix filter', () => {
    it('keeps rows whose path starts with any prefix', () => {
      const out = filterFolderActivity(sample, ['Active/'], new Set())
      expect(out.map((r) => r.folderPath)).toEqual(['Active/Foo', 'Active/Bar'])
    })

    it('supports multiple prefixes (union)', () => {
      const out = filterFolderActivity(sample, ['Active/', 'Archive/'], new Set())
      expect(out.map((r) => r.folderPath)).toEqual([
        'Active/Foo',
        'Active/Bar',
        'Archive/Old',
      ])
    })

    it('ignores empty-string entries within prefixes (would match everything)', () => {
      const out = filterFolderActivity(sample, ['', 'Active/'], new Set())
      // Empty prefix shouldn't match; only Active/* should survive.
      expect(out.map((r) => r.folderPath)).toEqual(['Active/Foo', 'Active/Bar'])
    })

    it('prefix without trailing slash matches the folder itself + descendants', () => {
      const withSelf: FolderActivity[] = [...sample, row('Active'), row('Active2')]
      const out = filterFolderActivity(withSelf, ['Active'], new Set())
      // Active matches; Active/Foo + Active/Bar match (descendants);
      // Active2 does NOT match (would have under the old startsWith).
      expect(out.map((r) => r.folderPath).sort()).toEqual(
        ['Active', 'Active/Bar', 'Active/Foo'].sort(),
      )
    })

    it('prefix with trailing slash matches descendants only (legacy form still works)', () => {
      const withSelf: FolderActivity[] = [...sample, row('Active')]
      const out = filterFolderActivity(withSelf, ['Active/'], new Set())
      // `Active/` legacy form: descendants only, not Active itself.
      expect(out.map((r) => r.folderPath)).toEqual(['Active/Foo', 'Active/Bar'])
    })

    it('prefix does not false-match sibling folder with same string prefix', () => {
      const siblings: FolderActivity[] = [row('案件'), row('案件2'), row('案件/A')]
      const out = filterFolderActivity(siblings, ['案件'], new Set())
      expect(out.map((r) => r.folderPath).sort()).toEqual(['案件', '案件/A'].sort())
    })
  })

  describe('leafNames filter', () => {
    it('keeps rows whose leaf segment matches', () => {
      const out = filterFolderActivity(sample, [], new Set(['重要客戶']))
      expect(out.map((r) => r.folderPath)).toEqual([
        'Inbox/SubA/重要客戶',
        'Inbox/SubB/重要客戶',
      ])
    })

    it('matches leaf only, not intermediate segments', () => {
      const out = filterFolderActivity(sample, [], new Set(['Inbox']))
      // No rows END with "Inbox", so nothing should match.
      expect(out).toEqual([])
    })
  })

  describe('combined filters (union semantics)', () => {
    it('returns rows matching prefix OR leaf', () => {
      const out = filterFolderActivity(
        sample,
        ['Archive/'],
        new Set(['重要客戶']),
      )
      expect(out.map((r) => r.folderPath)).toEqual([
        'Archive/Old',
        'Inbox/SubA/重要客戶',
        'Inbox/SubB/重要客戶',
      ])
    })
  })
})
