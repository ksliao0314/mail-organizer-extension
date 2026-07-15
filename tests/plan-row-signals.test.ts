import { describe, expect, it } from 'vitest'
import { aiOverrideInfo, willLearnFromOverride } from '@/popup/components/PlanRow'
import type { PlanItem } from '@/shared/types'

function item(over: Partial<PlanItem>): PlanItem {
  return {
    emailId: 'e1',
    emailSubject: 's',
    emailFrom: 'a@b.com',
    action: 'move',
    confidence: 0.9,
    reason: '',
    source: 'ai',
    ...over,
  }
}

// aiOverrideInfo mirrors the background wasUserOverride: an override is action
// change OR final target differing from the AI's original target.
describe('aiOverrideInfo', () => {
  it('no aiOriginalAction → not an override', () => {
    expect(aiOverrideInfo(item({})).overrode).toBe(false)
  })

  it('same action + same target → not an override', () => {
    const r = aiOverrideInfo(
      item({
        action: 'move',
        targetFolderPath: '03/Y',
        aiOriginalAction: 'move',
        aiOriginalTargetFolderPath: '03/Y',
      }),
    )
    expect(r.overrode).toBe(false)
    expect(r.finalTarget).toBe('03/Y')
    expect(r.aiOriginalTarget).toBe('03/Y')
  })

  it('user changed the target folder → override', () => {
    const r = aiOverrideInfo(
      item({
        action: 'move',
        targetFolderPath: '03/Right',
        aiOriginalAction: 'move',
        aiOriginalTargetFolderPath: '03/Wrong',
      }),
    )
    expect(r.overrode).toBe(true)
    expect(r.finalTarget).toBe('03/Right')
    expect(r.aiOriginalTarget).toBe('03/Wrong')
  })

  it('user changed the action (move → delete) → override', () => {
    expect(
      aiOverrideInfo(
        item({ action: 'delete', targetFolderPath: undefined, aiOriginalAction: 'move' }),
      ).overrode,
    ).toBe(true)
  })

  it('resolves new_folder targets from parent + name', () => {
    const r = aiOverrideInfo(
      item({
        action: 'new_folder',
        suggestedParentPath: '03',
        suggestedFolderName: '新案',
        aiOriginalAction: 'move',
        aiOriginalTargetFolderPath: '03/舊',
      }),
    )
    expect(r.overrode).toBe(true)
    expect(r.finalTarget).toBe('03/新案')
  })
})

describe('willLearnFromOverride', () => {
  it('override landing on a concrete move target → will learn', () => {
    expect(
      willLearnFromOverride(
        item({
          action: 'move',
          targetFolderPath: '03/Right',
          aiOriginalAction: 'move',
          aiOriginalTargetFolderPath: '03/Wrong',
        }),
      ),
    ).toBe(true)
  })

  it('override to delete (no concrete target) → does NOT show 將學習', () => {
    // delete overrides do not mint move-rules, so the chip must not appear.
    expect(
      willLearnFromOverride(
        item({ action: 'delete', targetFolderPath: undefined, aiOriginalAction: 'move' }),
      ),
    ).toBe(false)
  })

  it('AI suggestion the user accepted (no override) → no chip', () => {
    expect(
      willLearnFromOverride(
        item({
          action: 'move',
          targetFolderPath: '03/Y',
          aiOriginalAction: 'move',
          aiOriginalTargetFolderPath: '03/Y',
        }),
      ),
    ).toBe(false)
  })

  it('plain rule hit (no aiOriginalAction) → no chip', () => {
    expect(
      willLearnFromOverride(item({ source: 'rule', action: 'move', targetFolderPath: '03/Y' })),
    ).toBe(false)
  })
})
