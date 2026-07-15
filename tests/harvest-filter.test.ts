import { describe, expect, it } from 'vitest'
import { filterHarvestCandidates, type HarvestScanItem } from '@/background/harvest'
import { buildRuleIndex, newRule } from '@/shared/rules'
import type { ThreadMemoryEntry } from '@/shared/types'
import type { MovedByUsEntry } from '@/shared/storage'

type Moved = MovedByUsEntry

const scan = (over: Partial<HarvestScanItem>): HarvestScanItem => ({
  id: 'm-' + Math.random().toString(36).slice(2),
  subject: '關於 112訴204 開庭事宜',
  from: 'clerk@court.example.tw',
  conversationId: 'conv-' + Math.random().toString(36).slice(2),
  folderId: 'fid-A',
  folderPath: '案件/甲',
  ...over,
})
const moved = (over: Partial<Moved>): Moved => ({
  newMessageId: 'm1',
  destinationFolderId: 'fid-A',
  conversationId: 'conv1',
  movedAt: Date.now(),
  ...over,
})
const memEntry = (folderId: string): ThreadMemoryEntry => ({
  folderId,
  folderPath: 'p',
  lastFiledAt: '2026-01-01T00:00:00Z',
  timesFiled: 1,
})
const emptyOpts = () => ({
  movedByUs: [] as MovedByUsEntry[],
  ruleIndex: buildRuleIndex([]),
  convMemory: {} as Record<string, ThreadMemoryEntry>,
  subjectMemory: {} as Record<string, ThreadMemoryEntry>,
})

describe('filterHarvestCandidates', () => {
  it('keeps a clean manual filing and normalizes a long subject', () => {
    const out = filterHarvestCandidates([scan({ id: 'x', conversationId: 'c1' })], emptyOpts())
    expect(out).toHaveLength(1)
    expect(out[0]!.folderId).toBe('fid-A')
    expect(out[0]!.conversationId).toBe('c1')
    expect(out[0]!.normalizedSubject).toBeTruthy()
  })

  it('(1) excludes our own move by destination-side id', () => {
    const out = filterHarvestCandidates([scan({ id: 'mine' })], {
      ...emptyOpts(),
      movedByUs: [moved({ newMessageId: 'mine' })],
    })
    expect(out).toHaveLength(0)
  })

  it('(1) excludes our own move by conversation+folder near-match', () => {
    const out = filterHarvestCandidates([scan({ id: 'other', conversationId: 'cX', folderId: 'fid-A' })], {
      ...emptyOpts(),
      movedByUs: [moved({ newMessageId: 'unrelated', conversationId: 'cX', destinationFolderId: 'fid-A' })],
    })
    expect(out).toHaveLength(0)
  })

  it('(2) excludes mail already covered by an enabled rule', () => {
    const rule = newRule({
      type: 'domain',
      signal: 'court.example.tw',
      targetFolderId: 'fid_long_enough_to_pass_id_validation_xxx',
      targetFolderPath: '案件/甲',
      confidence: 0.8,
      source: 'ai_confirmed',
    })
    const out = filterHarvestCandidates([scan({ from: 'clerk@court.example.tw' })], {
      ...emptyOpts(),
      ruleIndex: buildRuleIndex([rule]),
    })
    expect(out).toHaveLength(0)
  })

  it('(3) excludes a conversation already routed to the SAME folder', () => {
    const out = filterHarvestCandidates([scan({ conversationId: 'c1', folderId: 'fid-A' })], {
      ...emptyOpts(),
      convMemory: { c1: memEntry('fid-A') },
    })
    expect(out).toHaveLength(0)
  })

  it('(3) KEEPS a conversation routed to a DIFFERENT folder (new signal)', () => {
    const out = filterHarvestCandidates([scan({ conversationId: 'c1', folderId: 'fid-A' })], {
      ...emptyOpts(),
      convMemory: { c1: memEntry('fid-OTHER') },
    })
    expect(out).toHaveLength(1)
  })

  it('drops the normalizedSubject for a too-short subject but still yields a candidate', () => {
    const out = filterHarvestCandidates([scan({ subject: '嗨', conversationId: 'c1' })], emptyOpts())
    expect(out).toHaveLength(1)
    expect(out[0]!.normalizedSubject).toBeUndefined()
  })
})
