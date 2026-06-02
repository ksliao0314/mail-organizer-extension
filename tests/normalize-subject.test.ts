import { describe, expect, it } from 'vitest'
import { normalizeSubject } from '@/shared/normalize'

// Imports the REAL implementation (was inline-copied here before 2026-05-22,
// which silently let the real version diverge from what tests checked).

describe('normalizeSubject', () => {
  it('strips single RE prefix (case-insensitive)', () => {
    expect(normalizeSubject('RE: 報價詢問')).toBe('報價詢問')
    expect(normalizeSubject('re: foo')).toBe('foo')
    expect(normalizeSubject('Re: hello')).toBe('hello')
  })

  it('strips nested prefixes (RE: RE: FW:)', () => {
    expect(normalizeSubject('RE: RE: 客戶來信')).toBe('客戶來信')
    expect(normalizeSubject('RE: Fw: FWD: deep')).toBe('deep')
  })

  it('strips Chinese prefixes including 轉寄 / 回覆 / 轉發', () => {
    expect(normalizeSubject('轉寄: 三月會議')).toBe('三月會議')
    expect(normalizeSubject('回覆:審閱')).toBe('審閱')
    expect(normalizeSubject('轉發: 文件')).toBe('文件')
  })

  it('handles fullwidth colon 全形:', () => {
    expect(normalizeSubject('RE:報告')).toBe('報告')
    expect(normalizeSubject('回覆:審閱')).toBe('審閱')
  })

  it('collapses whitespace and lowercases', () => {
    expect(normalizeSubject('   Hello   World  ')).toBe('hello world')
    expect(normalizeSubject('FOO\tBAR')).toBe('foo bar')
  })

  it('returns empty for whitespace-only', () => {
    expect(normalizeSubject('   ')).toBe('')
    expect(normalizeSubject('RE:')).toBe('')
  })

  it('groups RE: foo, Fw: FOO, foo all to same key', () => {
    expect(normalizeSubject('RE: foo')).toBe(normalizeSubject('Fw: FOO'))
    expect(normalizeSubject('foo')).toBe(normalizeSubject('RE: foo'))
  })

  // ---- Outlook auto-prefix extensions (added 2026-05-22) ----------------
  //
  // Real-world Exchange / O365 deployments stamp various prefixes that
  // should normalize away so the "same thread" detection works across
  // them. Before this batch, only the basic RE / FW / 回覆 etc. were
  // stripped; these system prefixes broke thread continuity.
  describe('Outlook auto-response prefixes', () => {
    it('strips 自動回覆 / Out of Office / Automatic reply', () => {
      expect(normalizeSubject('自動回覆: 假期通知')).toBe('假期通知')
      expect(normalizeSubject('Out of Office: away')).toBe('away')
      expect(normalizeSubject('Automatic reply: vacation')).toBe('vacation')
      expect(normalizeSubject('Auto-Reply: hi')).toBe('hi')
    })

    it('handles auto-response stacked with RE', () => {
      expect(normalizeSubject('RE: 自動回覆: 詢問')).toBe('詢問')
      expect(normalizeSubject('自動回覆: RE: 詢問')).toBe('詢問')
    })
  })

  describe('admin-stamped bracketed prefixes', () => {
    it('strips [External] / [外部]', () => {
      expect(normalizeSubject('[External] 客戶來信')).toBe('客戶來信')
      expect(normalizeSubject('[外部] 詢問')).toBe('詢問')
    })

    it('strips [已讀回條] / [已讀回執]', () => {
      expect(normalizeSubject('[已讀回條] 文件確認')).toBe('文件確認')
      expect(normalizeSubject('[已讀回執] foo')).toBe('foo')
    })

    it('strips [Spam] / [Junk] / [Caution]', () => {
      expect(normalizeSubject('[SPAM] foo')).toBe('foo')
      expect(normalizeSubject('[caution] please review')).toBe('please review')
    })

    it('handles bracketed prefix stacked with RE', () => {
      expect(normalizeSubject('RE: [External] 詢問')).toBe('詢問')
      expect(normalizeSubject('[External] RE: 詢問')).toBe('詢問')
    })
  })

  describe('numbered reply markers like Re[2]:', () => {
    it('strips Re[2]: form', () => {
      expect(normalizeSubject('Re[2]: foo')).toBe('foo')
      expect(normalizeSubject('RE[3]: bar')).toBe('bar')
    })
  })

  describe('Reply / R prefix variants', () => {
    it('strips Reply: ', () => {
      expect(normalizeSubject('Reply: status')).toBe('status')
    })

    it('strips bare R: ', () => {
      expect(normalizeSubject('R: status')).toBe('status')
    })
  })
})
