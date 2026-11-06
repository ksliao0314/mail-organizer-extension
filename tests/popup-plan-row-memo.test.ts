// A 勾選 re-renders only the rows it concerns (2026-11-06 review U5).
//
// PlanRow got `otherSelectedCount = selected.size − (this row selected)` as a
// prop, so every selection change changed that prop on EVERY row and
// React.memo re-rendered the whole list — 500 rows per click in a big plan —
// although only an expanded row's 「全部套用到勾選 N 項」 button reads it. The
// count now travels by context to that button alone.
//
// PlanRow is wrapped in a counting memo: the wrapper re-renders exactly when
// PlanScreen hands a row new props.

import { createElement, memo } from 'react'
import type { Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import type { PlanItem } from '@/shared/types'
import type * as PlanRowModule from '@/popup/components/PlanRow'
import { buttonContaining, cleanup, click, mountApp, stubRuntime } from './helpers/ui-harness'

const renders = vi.hoisted(() => new Map<string, number>())
vi.mock('@/popup/components/PlanRow', async (importOriginal) => {
  const mod = await importOriginal<typeof PlanRowModule>()
  const Counting = memo(function CountingPlanRow(props: PlanRowModule.PlanRowProps) {
    renders.set(props.item.emailId, (renders.get(props.item.emailId) ?? 0) + 1)
    return createElement(mod.PlanRow, props)
  })
  return { ...mod, PlanRow: Counting }
})

let root: Root | null = null
afterEach(() => {
  cleanup(root)
  root = null
  renders.clear()
})

const items: PlanItem[] = Array.from({ length: 6 }, (_, i) => ({
  emailId: `e${i}`,
  emailSubject: `mail ${i}`,
  emailFrom: 'a@acme.example',
  action: 'skip',
  confidence: 0.9,
  reason: '',
  source: 'ai',
}))

async function renderPlan() {
  stubRuntime((m) => {
    switch (m.type) {
      case 'getStatus':
        return {
          ok: true,
          data: {
            owaConnected: true,
            tokenValid: true,
            apiKeyConfigured: true,
            model: 'claude-sonnet-4-6',
            folderCacheCount: 0,
            excludePrefixes: [],
            batchSize: 50,
            internalDomains: [],
            primaryRootPath: '',
          },
        }
      case 'getOnboardingState':
        return { ok: true, data: { needed: false } }
      case 'getFolderActivity':
        return { ok: true, data: { activity: [] } }
      case 'getExecuteState':
        return { ok: true, data: { state: null } }
      case 'getAiClassifyProgress':
        return { ok: true, data: null }
      case 'getPopupState':
        return {
          ok: true,
          data: {
            phase: 'plan',
            updatedAt: Date.now(),
            plan: { items, folderTree: [], usage: null, summary: { ruleHits: 0, aiHandled: 6 }, banner: null },
          },
        }
      case 'savePopupState':
        return { ok: true }
      default:
        return undefined
    }
  })
  const { default: App } = await import('@/popup/App')
  const r = await mountApp(App)
  root = r.root
  await vi.waitFor(() => expect(r.host.querySelector('[data-plan-row-index="5"]')).not.toBeNull())
  return r.host
}

const row = (host: HTMLElement, i: number) => host.querySelector<HTMLElement>(`[data-plan-row-index="${i}"]`)!
const checkbox = (host: HTMLElement, i: number) =>
  row(host, i).querySelector<HTMLInputElement>('input[type="checkbox"]')!

describe('selection changes and PlanRow memo', () => {
  it('勾選 one row re-renders that row only — not the whole list', async () => {
    const host = await renderPlan()
    renders.clear()

    click(checkbox(host, 1))
    await vi.waitFor(() => expect(checkbox(host, 1).checked).toBe(true))
    click(checkbox(host, 3))
    await vi.waitFor(() => expect(checkbox(host, 3).checked).toBe(true))

    expect(renders.get('e1')).toBeGreaterThan(0)
    expect(renders.get('e3')).toBeGreaterThan(0)
    for (const id of ['e0', 'e2', 'e4', 'e5']) expect(renders.get(id) ?? 0).toBe(0)
  })

  it('an expanded row still shows the live 「全部套用到勾選 N 項」 count', async () => {
    const host = await renderPlan()
    click(row(host, 0).querySelectorAll('button')[0]!) // expand row 0
    await vi.waitFor(() => expect(buttonContaining(row(host, 0), '收起')).toBeDefined())
    expect(buttonContaining(row(host, 0), '全部套用到勾選')).toBeUndefined()

    click(checkbox(host, 2))
    await vi.waitFor(() => expect(buttonContaining(row(host, 0), '全部套用到勾選 1 項')).toBeDefined())
    click(checkbox(host, 4))
    await vi.waitFor(() => expect(buttonContaining(row(host, 0), '全部套用到勾選 2 項')).toBeDefined())
    click(checkbox(host, 0)) // the expanded row itself is not one of the "others"
    await vi.waitFor(() => expect(checkbox(host, 0).checked).toBe(true))
    expect(buttonContaining(row(host, 0), '全部套用到勾選 2 項')).toBeDefined()
  })
})
