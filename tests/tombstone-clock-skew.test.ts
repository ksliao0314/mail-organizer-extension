// Tombstones must beat the deleted rule's own timestamps (2026-11-05).
//
// Sync applies a tombstone only to rule versions whose max(createdAt,
// editedAt) <= deletedAt. A machine whose clock runs behind the machine that
// last edited a rule would otherwise mint a tombstone that never suppresses
// that version, so the delete silently fails to propagate.

import { describe, expect, it } from 'vitest'
import { deleteRule, newRule, tombstoneDeletedAt } from '@/shared/rules'
import { getRuleTombstones, setRules } from '@/shared/storage'

describe('tombstoneDeletedAt', () => {
  const NOW = Date.parse('2026-11-05T00:00:00Z')
  const CREATED = '2026-01-01T00:00:00Z'

  it('is now when the rule is older', () => {
    expect(tombstoneDeletedAt({ createdAt: CREATED, editedAt: NOW - 60_000 }, NOW)).toBe(NOW)
  })

  it('is strictly after an editedAt stamped by a clock that runs ahead', () => {
    expect(tombstoneDeletedAt({ createdAt: CREATED, editedAt: NOW + 600_000 }, NOW)).toBe(NOW + 600_001)
  })

  it('is strictly after a future createdAt and tolerates junk dates', () => {
    const created = Date.parse('2030-01-01T00:00:00Z')
    expect(tombstoneDeletedAt({ createdAt: '2030-01-01T00:00:00Z' }, NOW)).toBe(created + 1)
    expect(tombstoneDeletedAt({ createdAt: '' }, NOW)).toBe(NOW)
  })

  it('deleteRule uses it: the tombstone outranks a skewed editedAt', async () => {
    const r = {
      ...newRule({
        type: 'domain',
        signal: 'client.example',
        targetFolderId: 'AAMkADcwMWM5ZTE4LWYwZWMtNGUxYS1hZGMwLTdhNGE3',
        targetFolderPath: '03/A',
        confidence: 0.7,
        source: 'user_manual',
      }),
      editedAt: Date.now() + 10 * 60_000, // edited on a machine 10 minutes ahead
    }
    await setRules([r])
    await deleteRule(r.id)
    const [t] = await getRuleTombstones()
    expect(t!.deletedAt).toBeGreaterThan(r.editedAt)
  })
})
