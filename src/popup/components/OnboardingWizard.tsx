// First-run onboarding wizard.
//
// Distinguishes between three kinds of "empty popup" states:
//
//   (a) Genuinely new user — no API key, no rules. Just set the API key
//       and they're done. Everything else (primary root, internal
//       domains, etc.) is optional config exposed in Options; the first
//       classify works without any of it because the pipeline already
//       scans the user's whole folder tree.
//
//   (b) "Already used on another machine, just hasn't synced here yet."
//       Cloud has data from a different machineId. Going the (a) path
//       would have them set up from scratch, create rules on this
//       machine, then collide / overwrite when sync eventually pulls.
//       So we OFFER to enable sync + pull first.
//
//   (c) Reinstalled / wiped: local backups (chrome.storage.local
//       syncBackups) still exist from before. Offer restore.
//
// Wizard flow:
//
//   [ Welcome ] ── 我是第一次使用 ─→ [ API key ] ─→ [ Done ]
//        │
//        ├── 我在另一台機器用過 ─→ [ Existing-user options ]
//        │                          (cloud / import / restore)
//        │
//        └── 「我先看看、稍後再設定」 ─→ dismiss
//
// All terminal states call `onComplete()` (dismisses wizard + closes overlay).

import { useEffect, useState } from 'react'
import { CheckCircle2, Cloud, KeyRound, Upload } from 'lucide-react'
import { Button } from '@/components/ui/button'
import { send } from '@/shared/send-message'

type OnboardingState = {
  needed: boolean
  reasons: {
    hasApiKey: boolean
    hasRules: boolean
    hasPrimaryRoot: boolean
    dismissed: boolean
  }
  cloud?: {
    ruleCount: number
    tombstoneCount: number
    updatedAt: string
    sourceMachineId: string
  }
  syncEnabled: boolean
  hasLocalBackups: boolean
  backupCount: number
}

type Step =
  | 'welcome'
  | 'existing-options'
  | 'new-api-key'
  | 'done'

export function OnboardingWizard({
  onComplete,
}: {
  onComplete: () => void
}) {
  const [state, setState] = useState<OnboardingState | null>(null)
  const [step, setStep] = useState<Step>('welcome')
  const [apiKey, setApiKey] = useState('')
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)

  useEffect(() => {
    void (async () => {
      const r = await send<OnboardingState>({ type: 'getOnboardingState' })
      if (r.ok && r.data) setState(r.data)
    })()
  }, [])

  async function dismissAndClose() {
    await send({
      type: 'setSettings',
      patch: { onboardingDismissed: true },
    })
    onComplete()
  }

  async function handleEnableSync() {
    setBusy(true)
    setError(null)
    const r = await send<{ action: string; pulled?: boolean; ruleCount?: number }>({
      type: 'enableSync',
    })
    setBusy(false)
    if (!r.ok) {
      setError(r.message || '啟用同步失敗')
      return
    }
    await dismissAndClose()
  }

  async function handleSaveApiKey() {
    if (!apiKey.startsWith('sk-ant-')) {
      setError('API key 應該以 sk-ant- 開頭')
      return
    }
    setBusy(true)
    setError(null)
    const r = await send({ type: 'setApiKey', key: apiKey })
    setBusy(false)
    if (!r.ok) {
      setError(r.message || '儲存失敗')
      return
    }
    // Skip primaryRootPath / other config — those are optional, the
    // first classify works against the whole tree by default. User can
    // tune in Options later if they want focused 近日活動 panel etc.
    setStep('done')
  }

  if (!state) {
    return (
      <div className="flex items-center justify-center p-6 text-xs text-muted-foreground">
        載入中…
      </div>
    )
  }

  return (
    <div className="absolute inset-0 z-50 flex flex-col bg-background overflow-y-auto">
      <div className="flex-1 p-4 space-y-3">
        {step === 'welcome' && (
          <WelcomeStep
            onPickNew={() => setStep('new-api-key')}
            onPickExisting={() => setStep('existing-options')}
            onSkip={() => void dismissAndClose()}
          />
        )}
        {step === 'existing-options' && (
          <ExistingOptionsStep
            cloud={state.cloud}
            syncEnabled={state.syncEnabled}
            hasLocalBackups={state.hasLocalBackups}
            backupCount={state.backupCount}
            busy={busy}
            error={error}
            onEnableSync={() => void handleEnableSync()}
            onOpenOptionsImport={() => {
              chrome.runtime.openOptionsPage()
            }}
            onSwitchToNew={() => {
              setError(null)
              setStep('new-api-key')
            }}
            onBack={() => {
              setError(null)
              setStep('welcome')
            }}
          />
        )}
        {step === 'new-api-key' && (
          <NewApiKeyStep
            apiKey={apiKey}
            onApiKeyChange={setApiKey}
            busy={busy}
            error={error}
            onSave={() => void handleSaveApiKey()}
            onBack={() => {
              setError(null)
              setStep('welcome')
            }}
          />
        )}
        {step === 'done' && <DoneStep onClose={() => void dismissAndClose()} />}
      </div>
    </div>
  )
}

// ---- Welcome -------------------------------------------------------------

function WelcomeStep({
  onPickNew,
  onPickExisting,
  onSkip,
}: {
  onPickNew: () => void
  onPickExisting: () => void
  onSkip: () => void
}) {
  return (
    <>
      <div className="space-y-1.5">
        <h2 className="text-base font-semibold">歡迎使用 Mail Organizer</h2>
        <p className="text-[12px] text-muted-foreground">
          先告訴我你的情況，我會引導正確的設定方式：
        </p>
      </div>
      <button
        type="button"
        onClick={onPickNew}
        className="w-full text-left rounded-md border border-border hover:border-foreground/40 hover:bg-accent/30 p-3 transition-colors group"
      >
        <div className="flex items-center gap-2">
          <KeyRound className="size-4 text-muted-foreground group-hover:text-foreground" />
          <span className="font-medium text-xs">我是第一次使用</span>
        </div>
        <p className="text-[11px] text-muted-foreground mt-1 pl-6">
          設定 Claude API key 和主要案件資料夾、然後開始歸類
        </p>
      </button>
      <button
        type="button"
        onClick={onPickExisting}
        className="w-full text-left rounded-md border border-border hover:border-foreground/40 hover:bg-accent/30 p-3 transition-colors group"
      >
        <div className="flex items-center gap-2">
          <Cloud className="size-4 text-muted-foreground group-hover:text-foreground" />
          <span className="font-medium text-xs">我在另一台機器已經用過</span>
        </div>
        <p className="text-[11px] text-muted-foreground mt-1 pl-6">
          先檢查雲端 / 備份 / 匯出檔、避免從零設定蓋掉舊資料
        </p>
      </button>
      <button
        type="button"
        onClick={onSkip}
        className="w-full text-[11px] text-muted-foreground hover:text-foreground underline underline-offset-2 pt-2"
      >
        先進主畫面看看、稍後再設定
      </button>
    </>
  )
}

// ---- Existing-user options ----------------------------------------------

function ExistingOptionsStep({
  cloud,
  syncEnabled,
  hasLocalBackups,
  backupCount,
  busy,
  error,
  onEnableSync,
  onOpenOptionsImport,
  onSwitchToNew,
  onBack,
}: {
  cloud?: { ruleCount: number; tombstoneCount: number; updatedAt: string; sourceMachineId: string }
  syncEnabled: boolean
  hasLocalBackups: boolean
  backupCount: number
  busy: boolean
  error: string | null
  onEnableSync: () => void
  onOpenOptionsImport: () => void
  onSwitchToNew: () => void
  onBack: () => void
}) {
  return (
    <>
      <div className="space-y-1.5">
        <button
          type="button"
          onClick={onBack}
          className="text-[10px] text-muted-foreground hover:text-foreground underline underline-offset-2"
        >
          ← 回上一步
        </button>
        <h2 className="text-base font-semibold">先恢復舊資料</h2>
        <p className="text-[12px] text-muted-foreground">
          可以從以下任一來源把規則拉回來：
        </p>
      </div>

      {/* Top recommendation: cloud sync, if cloud has another machine's data */}
      {cloud ? (
        <div className="rounded-md border border-emerald-200 bg-emerald-50 p-3 space-y-2">
          <div className="flex items-center gap-2 text-xs font-medium">
            <Cloud className="size-3.5 text-emerald-700" />
            偵測到雲端資料
          </div>
          <p className="text-[11px] text-emerald-900/80 pl-5">
            另一台機器（ID {cloud.sourceMachineId.slice(0, 8)}…）已推送
            <span className="font-medium"> {cloud.ruleCount} 條規則 + {cloud.tombstoneCount} 條墓碑</span>，
            最後寫入於 {new Date(cloud.updatedAt).toLocaleString('zh-TW')}。
          </p>
          <p className="text-[11px] text-emerald-900/80 pl-5">
            啟用同步並拉下、這台機器就會跟另一台一致。
          </p>
          {error && (
            <p className="text-[11px] text-red-700 pl-5">{error}</p>
          )}
          <Button
            size="sm"
            onClick={onEnableSync}
            disabled={busy}
            className="ml-5 bg-emerald-700 hover:bg-emerald-800 text-white"
          >
            {busy ? '處理中…' : '啟用同步並拉下雲端規則'}
          </Button>
        </div>
      ) : (
        <div className="rounded-md border border-amber-200 bg-amber-50 p-3 space-y-1">
          <div className="text-[11px] text-amber-900">
            {syncEnabled
              ? '雲端目前沒看到其他機器寫入的資料。'
              : '同步未啟用、無法從雲端拉取資料。'}
          </div>
          <p className="text-[10px] text-amber-900/70">
            可能原因：另一台機器還沒推送過、或兩台機器的瀏覽器帳號不同步（Edge 不會跟 Chrome 共用）。
          </p>
        </div>
      )}

      {/* Alternative: import from file */}
      <button
        type="button"
        onClick={onOpenOptionsImport}
        className="w-full text-left rounded-md border border-border hover:border-foreground/40 hover:bg-accent/30 p-3 transition-colors group"
      >
        <div className="flex items-center gap-2">
          <Upload className="size-4 text-muted-foreground group-hover:text-foreground" />
          <span className="font-medium text-xs">從規則匯出檔匯入</span>
        </div>
        <p className="text-[11px] text-muted-foreground mt-1 pl-6">
          如果你在另一台機器用「匯出」按鈕存了 .json 檔、在「設定 → 規則庫」匯入
        </p>
      </button>

      {/* Alternative: restore from local backups (this device) */}
      {hasLocalBackups && (
        <button
          type="button"
          onClick={onOpenOptionsImport}
          className="w-full text-left rounded-md border border-border hover:border-foreground/40 hover:bg-accent/30 p-3 transition-colors group"
        >
          <div className="flex items-center gap-2">
            <CheckCircle2 className="size-4 text-muted-foreground group-hover:text-foreground" />
            <span className="font-medium text-xs">
              從本機備份回復（有 {backupCount} 份快照）
            </span>
          </div>
          <p className="text-[11px] text-muted-foreground mt-1 pl-6">
            如果同步把資料蓋掉了、或先前手動清空又後悔、可以從快照回復
          </p>
        </button>
      )}

      <div className="pt-2 border-t border-border">
        <p className="text-[11px] text-muted-foreground mb-1.5">
          以上都不適用？
        </p>
        <button
          type="button"
          onClick={onSwitchToNew}
          className="text-[11px] text-muted-foreground hover:text-foreground underline underline-offset-2"
        >
          ← 改用「我是第一次使用」流程
        </button>
      </div>
    </>
  )
}

// ---- New user: API key --------------------------------------------------

function NewApiKeyStep({
  apiKey,
  onApiKeyChange,
  busy,
  error,
  onSave,
  onBack,
}: {
  apiKey: string
  onApiKeyChange: (v: string) => void
  busy: boolean
  error: string | null
  onSave: () => void
  onBack: () => void
}) {
  return (
    <>
      <div className="space-y-1.5">
        <button
          type="button"
          onClick={onBack}
          className="text-[10px] text-muted-foreground hover:text-foreground underline underline-offset-2"
        >
          ← 回上一步
        </button>
        <h2 className="text-base font-semibold flex items-center gap-2">
          <KeyRound className="size-4" />
          設定 Claude API key
        </h2>
        <p className="text-[12px] text-muted-foreground">
          擴充用 Claude 來理解信件主旨並分類。API key 只存在這台機器、不會跟著同步離開。
        </p>
      </div>
      <div className="rounded-md border border-border p-3 space-y-2">
        <label className="text-[11px] text-muted-foreground block">
          貼上 API key（以 <code className="text-foreground">sk-ant-</code> 開頭）
        </label>
        <input
          type="password"
          value={apiKey}
          onChange={(e) => onApiKeyChange(e.target.value)}
          placeholder="sk-ant-..."
          className="w-full font-mono text-xs px-2 py-1.5 rounded border border-border bg-background focus:border-foreground/50 outline-none"
        />
        <p className="text-[10px] text-muted-foreground">
          沒有 key？前往{' '}
          <a
            href="https://console.anthropic.com/settings/keys"
            target="_blank"
            rel="noopener noreferrer"
            className="underline underline-offset-2"
          >
            console.anthropic.com
          </a>{' '}
          建立。
        </p>
      </div>
      {error && <p className="text-[11px] text-red-700">{error}</p>}
      <Button onClick={onSave} disabled={busy || apiKey.length < 10} className="w-full">
        {busy ? '儲存中…' : '儲存並繼續'}
      </Button>
    </>
  )
}

// ---- Done ---------------------------------------------------------------

function DoneStep({ onClose }: { onClose: () => void }) {
  return (
    <>
      <div className="space-y-1.5 pt-4">
        <div className="flex items-center gap-2">
          <CheckCircle2 className="size-5 text-emerald-700" />
          <h2 className="text-base font-semibold">設定完成</h2>
        </div>
        <p className="text-[12px] text-muted-foreground">
          現在可以按主畫面的「開始歸類」試跑一次。第一次跑會掃整個資料夾樹，前幾批 Claude 會學你的習慣，之後越來越多信會走規則自動歸（不用 token），AI 只在不確定時介入。
        </p>
      </div>
      <div className="rounded-md border border-emerald-200 bg-emerald-50 p-3 space-y-1">
        <p className="text-[11px] text-emerald-900">之後可考慮：</p>
        <ul className="text-[11px] text-emerald-900/80 list-disc list-inside space-y-0.5">
          <li>跑一次小批次（10-25 封）熟悉介面</li>
          <li>「設定 → 歸類偏好」可以指定主要根資料夾，讓「近日活動」面板聚焦在那個範圍</li>
          <li>「設定 → 初始掃描」一次性生成 domain 規則（可選）</li>
          <li>如果有第二台機器，「設定 → 跨機器同步」啟用</li>
        </ul>
      </div>
      <Button onClick={onClose} className="w-full">
        進入主畫面
      </Button>
    </>
  )
}
