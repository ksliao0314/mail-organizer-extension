// 「捨棄暫存計畫並重新載入」 for the popup's root error boundary (2026-11-06).
//
// A reload restores the saved plan (popupState) and re-merges the current
// classify run's AI progress — so when THAT data is what crashed the screen,
// every reload crashed again. This drops both first:
//   - clearPopupState: the plan snapshot (the batch's unexecuted edits);
//   - cancelClassify WITHOUT a runId, on purpose: the reload would adopt the
//     current run's progress, whichever surface started it, so that is the
//     data that has to go — and the user explicitly asked to discard. (The
//     DoneScreen's 回主畫面 cancels by id instead: there nothing is broken.)
//   - the FINISHED batch's execute state + its undo window (2026-11-06): the
//     popup reopens onto a finished batch's DoneScreen before it ever looks
//     at the plan, so a crash in the done / executing screens looped just the
//     same. `unlessRunning`: the SW refuses while a batch is executing — a
//     running batch is never discarded (its state is the batch's progress),
//     and then its undo snapshot is left alone too.

import { send } from '@/shared/send-message'

async function discardFinishedBatch(): Promise<void> {
  const r = await send<{ cleared?: boolean }>({ type: 'clearExecuteState', unlessRunning: true })
  if (r?.ok && r.data?.cleared !== false) await send({ type: 'dismissUndo' })
}

export async function discardSavedPlan(): Promise<void> {
  await Promise.allSettled([
    send({ type: 'clearPopupState' }),
    send({ type: 'cancelClassify' }),
    discardFinishedBatch(),
  ])
}
