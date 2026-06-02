import { describe, expect, it } from 'vitest'
import {
  browserLabel,
  detectBrowser,
  syncAccountDescription,
  syncSettingsUrl,
} from '@/shared/browser-detect'

function withUserAgent(ua: string, fn: () => void): void {
  const original = Object.getOwnPropertyDescriptor(navigator, 'userAgent')
  Object.defineProperty(navigator, 'userAgent', { value: ua, configurable: true })
  try {
    fn()
  } finally {
    if (original) Object.defineProperty(navigator, 'userAgent', original)
  }
}

describe('detectBrowser', () => {
  // Edge UA includes both "Edg/" and "Chrome/" — Edge must be checked first.
  it('detects Edge from "Edg/" token', () => {
    withUserAgent(
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 ' +
        '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0',
      () => {
        expect(detectBrowser()).toBe('edge')
      },
    )
  })

  it('detects Chrome (without Edg/) ', () => {
    withUserAgent(
      'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 ' +
        '(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      () => {
        expect(detectBrowser()).toBe('chrome')
      },
    )
  })

  it('detects Firefox', () => {
    withUserAgent(
      'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.15; rv:120.0) Gecko/20100101 Firefox/120.0',
      () => {
        expect(detectBrowser()).toBe('firefox')
      },
    )
  })

  it('returns "other" for unknown UA', () => {
    withUserAgent('Mozilla/5.0 (Mobile; rv:120.0) Goose/1.0', () => {
      expect(detectBrowser()).toBe('other')
    })
  })
})

describe('syncAccountDescription', () => {
  it('Edge → Microsoft account', () => {
    expect(syncAccountDescription('edge')).toContain('Microsoft')
  })
  it('Chrome → Google account', () => {
    expect(syncAccountDescription('chrome')).toContain('Google')
  })
  it('Firefox → Firefox Account', () => {
    expect(syncAccountDescription('firefox')).toContain('Firefox')
  })
})

describe('syncSettingsUrl', () => {
  it('Edge → edge://settings/...', () => {
    expect(syncSettingsUrl('edge')).toMatch(/^edge:\/\//)
  })
  it('Chrome → chrome://settings/...', () => {
    expect(syncSettingsUrl('chrome')).toMatch(/^chrome:\/\//)
  })
  it('Firefox → empty (no canonical settings URL we can reliably link)', () => {
    expect(syncSettingsUrl('firefox')).toBe('')
  })
})

describe('browserLabel', () => {
  it('returns short display names', () => {
    expect(browserLabel('edge')).toBe('Edge')
    expect(browserLabel('chrome')).toBe('Chrome')
    expect(browserLabel('firefox')).toBe('Firefox')
    expect(browserLabel('other')).toBe('瀏覽器')
  })
})
