# Mail Organizer — Claude Code 專案說明

Chrome MV3 擴充工具,把 Outlook 收件夾用「rule + Claude AI」兩段式歸類。權威使用者文件:`README.md`。本檔是給 Claude Code(及其他 AI 助手)接手時看的 architecture / convention / gotcha 速查。

最後一次大改:**2026-05-27**(規則庫重新設計、wipeMarker、a11y polish)。

---

## 核心架構

```
popup (React)        options (React)         content (outlook.*)
   │                    │                       │
   │  chrome.runtime.sendMessage                │
   ▼                    ▼                       ▼
              service-worker.ts (MV3 background)
                       │
       ┌───────────────┼──────────────┬──────────────┐
       │               │              │              │
   classifier.ts    execute.ts    initial-scan.ts  sync-engine.ts
   (Claude API)    (move 邏輯)   (生規則)         (跨機器同步)
       │
       ▼
   chrome.storage.local (持久) / .session (暫存) / .sync (跨機器)
```

代碼大小快查:`service-worker.ts` ≈ 2300 行、`options/App.tsx` ≈ 4400 行、`popup/App.tsx` ≈ 2000 行、`sync-engine.ts` ≈ 1100 行、`shared/rules.ts` ≈ 1300 行、`background/execute.ts` ≈ 1400 行。395 tests across 22 test files。

---

## 分類管線(5 層 fallback)

每封信走以下順序、先命中先勝:

1. **Thread Memory** — `service-worker.classifyPreflight` Step 0。同 ConversationId 或同 normalized subject 之前歸過、直接套用(信心 0.95 / 0.85)
2. **Rule Index** — `rules.matchEmailWithIndex`。5 種 type,**嚴格依 `TYPE_PRIORITY` 順序**迭代,先命中先 return:
   - 1: `case_code`(Latin 案件代號,主旨不分大小寫比對)
   - 2: `compound`(多條件 AND,例:domain + 主旨關鍵字)
   - 3: `domain`(任一信件地址匹配)
   - 4: `subject_keyword`(主旨子字串,bucket 內 longest-signal-first 排序)
   - 5: `sender`(完整 email,只在 From 命中時 trigger)
3. **Claude AI** — `classifier.classifyBatch`。對 unmatched 跑 LLM
4. **Threshold gate** — AI 信心 < `aiConfidenceThreshold`(預設 0.5) → 標 `source: 'unresolved'`、不執行
5. **Learning** — 執行成功後 `execute.generateAiConfirmedRules` / `generateAiOverrideRules` 自動產生規則

`ruleBeatsThread(rule)` 決定 thread memory 是否可被覆蓋:`case_code` / `compound` / `user_manual` 在 thread memory 之前;`domain` / `subject_keyword` / `sender` / 其他 source 之後。

---

## 規則學習:先廣後窄(2026-05-27 核心設計)

`background/execute.ts:chooseLearningSignal` 是學習主入口。哲學:

> 同網域第一次出現 → 建 plain-domain 規則(廣)。
> 第二次同網域進**另一個**資料夾 → 衝突偵測觸發 → 升級成 compound(網域 + 整段主旨,窄),原 broad 規則自動沉默。

### 6 個優先順序(由窄到廣)

| 優先 | 條件 | 學成 | 信心 | featureKind |
|---|---|---|---|---|
| 1 | 外部域 + 法院字號(`112訴204`) | `compound`(網域 + 字號) | 0.9 | `court_case` |
| 2 | 法院字號、無外部域 | `subject_keyword`(字號) | 0.85 | `court_case` |
| 3 | Latin 案件代號(`25A0067A`) | `case_code` | 0.9 | — |
| 4 | 可用網域(非 generic provider) + (條件式) | 預設 `domain`、衝突時 `compound`(網域 + 整段主旨) | 0.7 / 0.85 | `domain` / `full_subject` |
| 5 | 通用 provider(gmail.com / yahoo.com 等) + 完整地址 | `sender` | 0.7 | — |
| 6 | 內部網域 fallback(internalDomains 含網域) + 主旨可用 | `subject_keyword`(整段主旨) | 0.85 | `full_subject` |

**外部 vs 內部網域**:`internalDomainSet` 由 caller 傳入(`settings.internalDomains`)。空 set = solo 模式、所有外部網域可用。

**衝突偵測**:`existingRules.some(r => r.type==='domain' && normalizeSignal('domain',r.signal)===usableDomain && r.enabled && !r.orphaned && r.targetFolderPath !== item.targetFolderPath)`。重點:用 `normalizeSignal` 比對(處理 `@kgi.com` 等帶 `@` 前綴的 legacy 形式)。

**`extractSubjectSignal`**:把整段主旨正規化(去 Re:/Fwd:、空白合併、轉小寫)、當 signal。**不再 tokenize**(2026-05-27 之前曾走「拆詞 + lenient extractor」路徑,後改為主旨即特徵)。最短長度 `MIN_SUBJECT_SIGNAL_LEN = 3`。

`demoteOnly` flag:`chooseLearningSignal` 偵測到衝突但 subject 抽不出可用 signal 時、`item.aiOriginalAction` 還在、降為「只 demote 原衝突 broad 規則、不創新規則」。

---

## 規則 Confidence 上限(applyConfidenceCap)

`shared/rules.ts:applyConfidenceCap` 在**所有寫入點**強制 cap。`source==='user_manual'` 例外、其餘按 type:

| Type | Cap |
|---|---|
| `domain` | 0.7 |
| `sender` | 0.75 |
| `subject_keyword` | 0.9 |
| `compound` / `case_code` | 0.95 |

**寫入點覆蓋**(2026-05-13 之後):
- `upsertRule` — Options 編輯
- `addRules` — 任何批次新增
- `bumpRuleHits` — 升信心也卡 cap
- `sync-engine.doPull` — cloud 規則進來前 cap(防舊版機器推高信心域名規則)
- `sync-engine.restoreBackup` — 從 backup 還原也 cap
- `service-worker.importRules` — 匯入時 cap,且 surface `cappedCount` + `cappedMaxDelta` 給 UI

---

## Storage Keys

### `chrome.storage.local`(持久)

| Key | 內容 | Cap / TTL |
|---|---|---|
| `settings` | 所有設定(含 API key) | — |
| `rules` | 規則陣列 | — |
| `ruleHistory` | 規則編輯 audit log | 500 events |
| `ruleTombstones` | 已刪除規則(防自動復活) | 2000 本機 / 500 同步 |
| `skipHistory` | 使用者選擇保留的 email IDs | 5000 + 60 天 TTL |
| `folderCache` | Outlook 資料夾樹快取 | — |
| `folderActivity` | 近日活動 ledger | 200 |
| `folderActivityRefreshAt` | popup auto-refresh 節流 | — |
| `conversationMemory` | ConvId → folder map | 5000 |
| `subjectMemory` | normalized subject → folder + conflictCount | 3000 |
| `metrics` | 累積統計 | — |
| `undoSnapshot` | 30 秒撤回快照 | 過期清除 |
| `weeklyDigest` | 週報 snapshot | — |
| `syncBackups` | 同步前的本機 snapshot rotation | 5 |
| `syncLastError` | 上次同步失敗紀錄 | next success 清空 |
| `remoteWipeNotice` | 收到另一台機器 wipeMarker 後的本地通知 | 用戶 dismiss 清除 |
| `errorLog` | 中央 SW 錯誤集中 | 200 entries |

### `chrome.storage.session`(SW 生命週期)

`cachedToken` / `aiClassifyProgress` / `popupState` / `preflightCache` / `classifyStage`

### `chrome.storage.sync`(跨機器,100 KB quota)

| Key | 內容 |
|---|---|
| `syncMeta` | schemaVersion / sourceMachineId / updatedAt / chunkCounts / `recentPushes[]` / **`wipeMarker?`** |
| `syncSettings` | settings minus per-device fields |
| `syncRules_0..N` | user_manual + ai_confirmed + ai_overridden 規則 |
| `syncTombstones_0..M` | 最近 500 條墓碑 |
| `syncFolderActivity_0..K` | 最近 20 個 folder activity(v2 schema) |

---

## 跨機器同步(sync-engine.ts)

### 同步什麼 / 不同步什麼

**同步**(`shouldSyncRule(r)` returns true):
- `source` ∈ {`user_manual`, `ai_confirmed`, `ai_overridden`}
- `!orphaned`

**不同步**:
- `auto_scan` 規則(每台機器跑自己的初始掃描)
- orphaned 規則
- `claudeApiKey` / `syncMachineId` / `lastSyncAt` / `syncEnabled` / `onboardingDismissed`(`PER_DEVICE_SETTINGS_FIELDS`)
- `folderCache` / `skipHistory` / `conversationMemory` / `subjectMemory`(本機快取)

### 觸發鏈

```
LOCAL 變更 → chrome.storage.onChanged → installSyncListener
  if changes['rules' | 'ruleTombstones' | 'folderActivity'] OR user-facing settings
    → schedulePush('local-mutation')  ← 5s debounce coalesces
      → doPush()

REMOTE syncMeta 變更 → listener
  if newMeta.sourceMachineId !== ours
    → handleRemoteSyncMetaChange  ← 200ms debounce coalesces bursts (P-2)
      → dispatchRemotePull
        → doPull('remote-change', { mode: 'union' })
          ← union-with-tombstones merge (Bug #E)
          ← wipeMarker check + auto-apply (2026-05-27)
```

### 兩個 pull flag(Bug #G + #O 兩階段修法)

```ts
let pullActive = false      // writes 期間 true、同步 set 在 await 之前、防併發 doPull
let pullInProgress = false  // pull 起到結束後 1s grace、防 macrotask listener 觸發 echo
let pullGraceTimer = null   // setTimeout 在 finally 排程、到期才清 pullInProgress
```

**Iron law**: `pullActive` check + set **必須同步**(無 await 之間)。否則兩個 doPull 都會通過 check 進到實作(TOCTOU)。

### Per-device 欄位永不離開本機

```ts
// sync-engine.ts:PER_DEVICE_SETTINGS_FIELDS
const PER_DEVICE_SETTINGS_FIELDS = new Set([
  'claudeApiKey',          // 安全
  'syncMachineId',         // 識別「我們是誰」、別機器寫過來 = 身分汙染
  'lastSyncAt',            // 各機器自己的 sync 時鐘
  'syncEnabled',           // 各機器自己的開關
  'onboardingDismissed',   // 各機器自己的 first-run 狀態(Bug #T)
])
```

`push: stripPerDeviceSettings → 雲端只看到 safe fields`
`pull: filter incomingSettings → 本機 per-device 不被覆蓋`
`restoreBackup: 同 filter → 即使 backup 含這些欄位也不還原`

### 跨機器 wipe 傳遞(wipeMarker,2026-05-27)

問題:union-mode pull 會保留本機沒在 cloud 出現的同步規則。在 A 機按「全部刪除」後,B 機下一次推送會把它的舊規則送進 cloud,A 機 pull 回來、wipe 等於白做。

機制:
1. `wipeAllRules` handler:
   - 先 quiesce + clearCloudState(失敗則 **abort 前不動本機**——避免 tombstone resurrection)
   - 本機 `setRules([])` / clearAllRuleTombstones / clearAllAiMemory / clearRuleHistory
   - `pushNow('post-wipe', { wipeMarker: true })` 在 syncMeta 刻 `wipeMarker: { at, byMachineId }`
2. 後續 normal push 會**保留** cloud 現有的 wipeMarker(只有新的 wipe 會覆蓋)
3. 其他機器 `doPull` 偵測 `meta.wipeMarker.byMachineId !== ours && settings.lastSyncAt !== '' && meta.wipeMarker.at > settings.lastSyncAt`:
   - 在 pullActive critical section 內 mutateRules → 過濾 `shouldSyncRule` 的規則(**保留 tombstones**,F3:本機刪除意圖不因遠端 wipe 而消失)
   - **套用後立刻 `setSettings({ lastSyncAt: marker.at })`**(F2:在 chunk 讀取 / merge 之前,避免下游拋錯導致 marker 反覆 re-trip)
   - 持久化 `remoteWipeNotice` 到 `chrome.storage.local`、UI 浮 banner 顯示「另一台機器執行了全部刪除、本機已套用、可從備份回滾」
4. 三道 guard:
   - echo guard(`byMachineId !== ours`)— 不觸發自己的 wipe
   - **first-participant guard(`lastSyncAt !== ''`,F1)** — 從未同步過的機器豁免;marker 只對「已在同步集裡」的機器生效,首次啟用同步(union pull)不被 cloud 殘留的舊 wipe 清空本地規則庫
   - staleness guard(`marker.at > lastSyncAt`)— 不重複觸發已套用的 marker

對應 API:`readRemoteWipeNotice()` / `dismissRemoteWipeNotice()` / message types `getRemoteWipeNotice` / `dismissRemoteWipeNotice`。

### Schema 升級

`SYNC_SCHEMA_VERSION = 2`。

- v1 → v2 (2026-05-26):加 folderActivity 同步 + recentPushes
- v1 client 拉 v2 cloud:**refuse + 紅 banner + 提示升級**
- v2 client 拉 v1 cloud:**OK**(folderActivityChunkCount 預設 0)
- v2 push 蓋 v1 cloud:**OK**(升級)
- v1 push 想蓋 v2 cloud:**被 Bug #H downgrade protection 擋下**

### 已修同步 Bug 一覽

```
#A 雲端 per-device 欄位汙染          Bug
#B first-enable 丟本機規則            Critical
#C restore 蓋寫 machineId             Bug
#D 規則超過 chunk cap 沒提示          UX
#E remote-pull replace 丟本機規則     Critical
#F restore 沒 reset tombstones        Bug
#G pullInProgress race                Critical
#H downgrade attack                   High
#I setSettings 無 mutex               Race
#J pull 錯誤吞 console                UX
#K backup 10 個吃 quota               Storage
#L refresh 全失敗仍 stamp 時間        Bug
#M enableSync 顯示假成功              UX
#N disable+clearCloud race            Bug
#O concurrent doPull                  Race
#P icon 開重複 OWA tab                UX
#Q appendBackup RMW                   Race
#T onboardingDismissed 應 per-device  Privacy
#U logError 無 mutex                  Race
#V recentPushes 沒防 malformed payload Defensive
P-1 writeFolderActivity 無腦寫        Perf
P-2 remote-pull 無 debounce           Perf
P0 clearCloud race + wipeMarker       Critical (2026-05-27)
P0 tombstone resurrection on fail     Critical (2026-05-27)
```

每個都有對應 regression test 在 `tests/sync-engine.test.ts` 或 `tests/sync-wipe-marker.test.ts`。

---

## 規則生命週期(2026-05-27 完整收口)

| 階段 | 觸發 | 行為 |
|---|---|---|
| **自動生** | 執行成功後 `generateAiConfirmedRules` / `generateAiOverrideRules` | 走 `chooseLearningSignal` 6 層、`applyConfidenceCap`、`addRulesFilteringTombstones` |
| **自動沉寂** | `stale-sweep.ts:autoDisableStaleRules`(每 6 小時 + 啟動跑一次) | `matchCount ≥ 20` 且 `overrideCount/matchCount ≥ 0.5` → 設 `enabled=false` + `autoDisabledReason: 'high-error-rate'` + `autoDisabledAt: now` |
| **stale 直接刪** | 同上 sweep | `matchCount === 0` 且 createdAt 超過 100 天 OR `matchCount > 0` 且 lastUsedAt 超過 100 天 → **完全刪除**(不留 tombstone、不寫 audit、不 console)。設計理由:長期沒命中等於不存在,留痕跡只汙染信任 |
| **legacy token 清理** | 同上 sweep | 2026-05-27 之前用 tokenize 路徑生的 subject_keyword 規則(`isLegacyTokenSubjectRule`):createdAt < `REDESIGN_CUTOFF_ISO` 的 ai_confirmed / ai_overridden subject_keyword 規則 → 設 `autoDisabledReason: 'legacy_token'` |
| **自動復活** | `addRulesFilteringTombstones` dedup | 同 type+signal+target 對應到一條 auto-disabled 規則 → 自動 `enabled=true` + 清 `autoDisabledAt` / `autoDisabledReason` |
| **使用者手動 toggle** | UI 按鈕 | 清掉 auto-disable 痕跡(避免 UI 視覺錯誤) |
| **永不自動觸碰** | — | `source==='user_manual'`(神聖) |
| **不會自動學的** | — | `internalDomains` 內網域的 sender 規則(內部 sender 通常太雜)、generic provider 純 domain 規則(會學成 sender) |

`dedupeRulesByKey`(2026-05-27):啟動時跑一次,把 (type, normalizeSignal, target) 三元組相同的規則合併為一個(取最高 matchCount)。修復 import + sync 歷史殘留的重複規則。

---

## 重點 Conventions

- **規則寫入永遠走 mutex** — `mutateRules` 包裹所有寫入。`addRulesFilteringTombstones` mutex 內 dedup、防並發 race
- **目標資料夾 ID 用 sanitize** — execute 寫規則前用 `sanitizeRuleTargetFolderId` 去掉 `pending:` 前綴、空字串、PLACEHOLDER 等
- **Tombstones 防復活** — 自動生成的規則(`auto_scan` / `ai_confirmed` / `ai_overridden`)永遠走 `addRulesFilteringTombstones`、user_manual 不受此限
- **`user_manual` 神聖不可侵** — auto 路徑永遠不會 disable / overwrite 使用者手寫規則
- **AI 路徑 fallback 到 unresolved** — classifier 拿到不存在的資料夾路徑(move / new_folder 的 suggestedParentPath)會降為 `source: 'unresolved'`、而非執行階段才爆
- **API key 永遠不外洩** — 不 log、不在錯誤訊息、diagnostic export `[redacted]`、settings export 排除、跨機器同步 strip
- **applyConfidenceCap 是寫入時 invariant** — 不要在讀取後修補,而是在所有寫入點(upsertRule / addRules / bumpRuleHits / sync pull / restoreBackup / importRules)強制

---

## 通用化承諾(2026-05-22 之後不可破)

**禁止硬編碼**:
- 任何特定事務所 / 公司名稱
- 個人 email、真人姓名
- 任何特定使用者的資料夾命名慣例(編號前綴、內部分類名)
- 「事務所」「律師」「案件」字眼(在 user-facing 文字中,改用中性詞)

**改用**:
- `settings.internalDomains` 取代 `INTERNAL_DOMAIN` 常數
- `settings.primaryRootPath` 取代 `'案件'`
- `settings.internalSubjectCategories` 取代固定關鍵字列表
- 「歸類偏好」「內部」「資料夾」 取代 「事務所設定」「同事」「案件」

新增 setting 時記得同步:`types.ts` Settings + `types.ts` DEFAULT_SETTINGS + `storage.ts` sanitize + `service-worker.ts` setSettings handler whitelist + getStatus payload + options/popup UI + **考慮是否 per-device**(若是,加進 `PER_DEVICE_SETTINGS_FIELDS`)。

---

## Options 頁結構(2026-05-27 重新設計)

3 個分頁(原 5 個):**分類引擎** / **規則庫** / **資料 & 同步**。

舊 hash redirect:`#connection` → `#engine`,`#skip` → `#data`。

### 規則庫頁兩種模式

1. **儀表板**(預設):4 卡片(總數 / 啟用 / 衝突 / 休眠)+ Top 5 命中規則 + 「打開規則庫」按鈕
2. **全螢幕工作區**(`#rules-library/all` 等):進 `RuleLibraryView`,左側 6 個子分頁:
   - `/all` — 全部規則,預設**按資料夾分組**(`RuleGroupedTable`),`viewMode` 可切 grouped / flat,collapsed 狀態存 sessionStorage
   - `/conflicts` — 衝突清單 + 一鍵自動升級為 compound
   - `/dormant` — 自動休眠的規則
   - `/health` — 規則健康度(orphaned / 高錯誤率 / 重複)
   - `/scan` — 初始掃描重跑入口
   - `/history` — `ruleHistory` audit log(最近 500)

### Drawer 與 a11y

`RuleDetailDrawer`(右側抽屜):
- `role="dialog" aria-modal="true"`
- 開啟自動 focus 關閉按鈕、關閉自動 restore focus 到觸發 row
- Tab 在 drawer 內 cycle、Tab 從最後一個 wrap 到第一、Shift+Tab 反向 wrap
- Esc 關閉(`stopImmediatePropagation` + `preventDefault` 防止 OWA FAB 攔截)
- backdrop `tabIndex={-1}` 不進 Tab cycle

`RuleRow`:`tabIndex={0}` + Enter/Space → 開 drawer、`aria-label` 完整描述、`focus:ring-2 focus:ring-foreground/20`

`RuleLibraryActionsMenu`:`role="menu"`、`role="menuitem"`、Arrow ↑↓ 導覽 / Home / End / Esc 關閉並 restore trigger focus

`ConfirmDialog`:`role="alertdialog"`、Esc 取消、backdrop click 取消、`autoFocus` 在 Cancel(避免 Enter 誤確認危險動作)

`RuleForm`:`useEffect([initial?.id])` reset state、避免切換規則時殘留前一條的值。同一條規則的同步更新**不會 clobber 進行中的編輯**(只有 id 變才 reset)

---

## Gotchas

- **Thread memory 不能蓋過 user_manual / case_code / compound** — `classifyPreflight` 用 `ruleBeatsThread` helper 排序
- **bumpRuleHits errorRate gate** — overrideCount/matchCount > 0.2 時不升信心
- **subject_keyword 用 longest-match** — `buildRuleIndex` 內 `subjectKeywordSort` 按 signal.length desc 排序
- **內部網域邏輯需顯式設定** — `internalDomains: []` 是 valid 狀態(solo / 個人用戶)、不可假設一定有值
- **migration 在 `getSettings` 內** — 舊使用者(storage 已有 settings 但缺 internalDomains 等欄位)自動補回原硬編碼值
- **chooseLearningSignal 接收 internalDomainSet 由 caller 傳入** — 不在裡面 await getSettings、避免 tight loop 多次 storage read
- **chooseLearningSignal 衝突偵測用 normalizeSignal** — 不是 `.toLowerCase()`,因為 import / manual 路徑可能存進 `@kgi.com` 帶 `@` 前綴的形式
- **autoUpgradeConflictRules 用 freshFolderPath** — 從 `folderById.get(effectiveFolderId).path` 取當下路徑,避免規則的 stale path 被烙進新升級的 compound 規則
- **Outlook DisplayName 含 `/` 編碼為 U+FF0F** — `encodeFolderName` / path 比對都用 fullwidth solidus
- **OData filter PascalCase** — `From/EmailAddress/Address`、不是小寫;`autoUpgradeConflictRules` 若 server filter 失敗會 fall back 到 client-side filter
- **wipeAllRules abort on cloud-clear failure** — 雲端清除失敗時不動本機,避免 cloud 滿、local 空的不一致狀態
- **REDESIGN_CUTOFF_ISO = '2026-05-27T00:00:00.000Z'** — `isLegacyTokenSubjectRule` 用這個分界線辨識舊 tokenize 路徑生的規則
- **非冪等 retry 是禁忌** — `outlook-api.ts:request()` 對 POST / DELETE / PATCH **不** retry 網路錯誤 + 5xx,只 retry 429。若硬把這條規則破掉,POST `/move` 第一次成功但 response 掉了的場景會在重試時 404、UI 顯示成失敗、使用者誤以為訊息沒搬走。`OutlookError.uncertain = true` 是給這類「可能成功」case 用的 marker
- **404 ErrorItemNotFound on move/delete = soft skip,不是 error** — `executeItem` 用 `isAlreadyMovedError(e)` helper 偵測、回傳 `status: 'skipped'`。原因:訊息已不在那個 Id(可能我們上一次 POST 其實成功了、或被其他規則 / 帳號處理過)。重試只會再 404

---

## 已驗證但易踩雷的設計

- **Pipeline 模式預跑** — DoneScreen 觸發、`prefetchStartedRef` 鎖住 phase 不會重跑、aiClassifyProgress 存活時跳過
- **AI invalid JSON 復原** — `parseAiActions` 三層:strict / balanced-array / loose(strip 註解 + 尾逗號 + smart quotes)
- **長 subject 截斷** — `buildEmailBlock` 內 `MAX_SUBJECT_LEN = 200`、保護 prompt budget
- **檔案匯入版本檢查** — settings export 帶 `version: 1`,import 比對 `SUPPORTED_IMPORT_VERSION` 並警告
- **規則匯入 confidence cap surface** — `importRules` 回傳 `cappedCount` + `cappedMaxDelta`,UI 顯示「N 條規則的信心度從 X 降至 Y(符合系統上限)」
- **`skipFlagged` 設定 hardcoded** — UI 移除、SW 永遠跳過 Outlook flagged 信件(`Flag.FlagStatus === 'Flagged'`)、這是預設值

---

## 規則類型 ↔ 學習路徑速查

| 信件特徵 | 規則 type 學成 | featureKind |
|---|---|---|
| 外部域 + 法院字號(`112訴204`) | `compound` | `court_case` |
| 法院字號、沒外部域 | `subject_keyword` | `court_case` |
| Latin 案件代號(`25A0067A`) | `case_code` | — |
| 外部域、無衝突 | `domain` | `domain` |
| 外部域、**有同網域指向不同 folder 的衝突** | `compound`(網域 + 整段主旨) | `full_subject` |
| 通用 provider(gmail.com)+ 完整地址 | `sender` | — |
| 內部網域(internalDomains 含) + 主旨可用 | `subject_keyword`(整段主旨) | `full_subject` |
| 找不到任何 signal | — | null(skip 學習) |

---

## 測試

`tests/` 目錄、vitest + jsdom + mocked chrome.storage。**395 tests across 22 files**。重要測試:

- `compound-learning.test.ts` — 學到內部網域 / compound 規則
- `case-code-learning.test.ts` — case_code 規則學習
- `ai-override.test.ts` — 使用者改 AI 後的 override flow
- `rules.test.ts` — match/index/normalize core + applyConfidenceCap
- `rule-health.test.ts` — orphan / conflict 偵測
- `rule-history.test.ts` — audit log shape
- `rule-io.test.ts` — export/import v2(含 tombstones)
- `rule-reconcile.test.ts` — folder 樹 reconcile
- `tombstones.test.ts` — tombstone CRUD + sync cap
- `classifier.test.ts` — prompt 組裝、JSON 解析
- `sync-engine.test.ts` — 41 個 sync regression tests(Bug A-V + P-1/P-2)
- `sync-wipe-marker.test.ts` — wipeMarker 跨機器傳遞(8 tests)
- `sync-chunks.test.ts` — chunking helpers
- `folder-activity-filter.test.ts` — 近日活動篩選(prefix with/without slash)
- `browser-detect.test.ts` — UA 偵測(Edge before Chrome!)
- `auto-propagate.test.ts` — 同 conversation 後續 inherit AI 決策
- `court-case-extract.test.ts` — `112訴204` 等台灣字號 regex
- `exemplars.test.ts` — `selectExemplars` few-shot 挑選
- `normalize-subject.test.ts` — Re:/Fwd:/Out of Office 等前綴剝除
- `path-encoding.test.ts` — U+FF0F fullwidth solidus
- `storage.test.ts` / `storage-sanitize.test.ts` — storage 層 mutex + sanitize

**Mock 限制**:`tests/setup.ts` 內的 `chrome.storage.onChanged` 只 record listeners、**不 fire 事件**。listener-driven 行為(sync echo loop / popup 自動 refresh)只能透過 contract 驗證。e2e 用 Playwright 載真 extension。

---

## MV3 / Chrome gotchas(踩過、寫下來)

- **SW idle-kill 30s** — `installKeepAliveListener` 用 `chrome.alarms` 撐住 long execute。喚醒時 `recoverStaleStates` 把 inProgress=true 但無 worker 的 task 標 failed
- **chrome.storage.onChanged 是 macrotask dispatch** — `set()` Promise resolved 之後、listener 才在下一個 task tick fire。Bug #G root cause;fix 靠 setTimeout 延後清 pullInProgress
- **chrome.action.setPopup 跨 SW 重啟持久化** — 上次 setPopup('') 重開 browser 仍有效。action-router 開機跑一次 syncPopupConfig 重新校正
- **chrome.tabs.onUpdated 看不到不在 host_permissions 的 URL** — navigating outlook → google 時 changeInfo.url 是 undefined。action-router 不 gate URL、全部 onUpdated 都 sync(debounced)
- **chrome.storage.sync 速率限制** — 100 KB total / 8 KB per item / 120 writes/min / 1800 writes/hour。我們 6KB chunks + 5s debounce → 遠低於上限
- **chrome.tabs.create 不需 tabs permission**;但 `chrome.tabs.query({url})` 需要 host_permissions 涵蓋
- **JS async TOCTOU** — single-thread but await 中可被插入。Bug #O root cause;guard set MUST 在 await 之前
- **`finally` 是同步** — try 內所有 await resolve 之後、finally 同步執行、再 return。但 finally 中 `setTimeout` 排的 callback 是 macrotask、後跑

---

## 目錄結構(2026-05-27 後)

```
mail-organizer-extension/
├── README.md                       ← 使用者文件(權威)
├── CLAUDE.md                       ← 本檔(AI 助手 / 開發者速查)
├── CHANGELOG.md                    ← 重大變更時間軸
├── manifest.config.ts              ← crxjs MV3 manifest
├── src/
│   ├── background/
│   │   ├── service-worker.ts       ← 主 dispatcher (~2300 行)
│   │   ├── sync-engine.ts          ← 跨機器同步 (~1100 行)
│   │   ├── action-router.ts        ← icon 點擊路由(無 OWA → 開分頁)
│   │   ├── execute.ts              ← 批次執行 + retry + undo + 學習 (~1400 行)
│   │   ├── initial-scan.ts         ← Pass A / A2 / B / C
│   │   ├── stale-sweep.ts          ← 每日休眠 / stale 刪除
│   │   ├── keep-alive.ts           ← 防 SW idle-kill
│   │   ├── token.ts                ← MSAL token cache
│   │   └── handlers/
│   │       └── sync.ts             ← 沙箱化的 sync handlers
│   ├── shared/
│   │   ├── constants.ts            ← 所有 magic numbers
│   │   ├── types.ts                ← Rule / Settings / FolderActivity / 等
│   │   ├── rules.ts                ← matchEmailWithIndex / mutateRules / applyConfidenceCap (~1300 行)
│   │   ├── storage.ts              ← chrome.storage.local + 全部 mutex
│   │   ├── sync-chunks.ts          ← chunking helpers
│   │   ├── classifier.ts           ← Claude API + prompt caching (extended-cache-ttl 1h)
│   │   ├── rule-io.ts              ← export/import v2(含 tombstones)
│   │   ├── rule-health.ts          ← orphan / conflict / overrideRate 偵測
│   │   ├── outlook-api.ts          ← Graph / Outlook REST 封裝
│   │   ├── error-log.ts            ← 中央 SW 錯誤紀錄
│   │   ├── folder-activity-filter.ts ← 近日活動篩選(prefix/leaf)
│   │   ├── browser-detect.ts       ← Edge / Chrome / Firefox 偵測
│   │   ├── normalize.ts            ← normalizeSubject + Re:/Fwd: 等
│   │   └── messages.ts             ← popup ↔ SW message 型別
│   ├── popup/
│   │   ├── App.tsx                 ← state machine + screens (~2000 行)
│   │   └── components/
│   │       ├── FolderPicker.tsx
│   │       ├── OnboardingWizard.tsx
│   │       └── PlanRow.tsx
│   ├── options/
│   │   └── App.tsx                 ← 3 tabs + 規則庫工作區 (~4400 行)
│   ├── content/
│   │   ├── content.ts              ← token 抓取
│   │   └── owa-fab.ts              ← 浮動按鈕注入
│   └── components/ui/              ← shadcn 基礎元件
└── tests/                          ← 395 tests
```

---

## 重要 invariants(NEVER break)

1. **claudeApiKey 永不離開本機** — 不入 sync、不入 backup、不入 export
2. **PER_DEVICE_SETTINGS_FIELDS 永不 round-trip** — push 時 strip、pull 時 filter、restore 時 filter
3. **user_manual rule 永不被自動停用 / 自動修改 / 自動刪除** — staleSweep 跳過 `source==='user_manual'`,applyConfidenceCap 也跳過,**`dedupeRulesByKey` comparator 把 user_manual 排在 enabled 偏好之前(F8:停用的 user_manual 不會輸給啟用的 auto 規則被硬刪)**
4. **Tombstones 一旦寫、AI/scan 不可重生同 triple** — `addRulesFilteringTombstones` 強制
5. **OWA folderId 不可硬編** — 透過 `folderCache` + reconcile
6. **TOCTOU guard 在 await 之前** — `pullActive` / 任何 in-flight flag 的 check+set 必須同步
7. **applyConfidenceCap 在所有寫入點** — 不要修補讀路徑、要 enforce 寫路徑
8. **規則 mutation 永遠走 mutateRules** — 包括 import / sync pull / reconcile / bumpRuleHits / **wipeAllRules(F12:用 `mutateRules(()=>({next:[]}))` 而非 raw `setRules([])`,否則並發 batch 寫入可在 wipe 後復活規則)** / etc. 不要再有 raw `setRules` 散在 handler 裡
9. **wipeMarker 寫入時不抹除舊 marker** — `doPush` 在 normal push 時保留 `cloudMeta.wipeMarker`、只有 wipe push 才覆蓋(讓第三台機器後加入也能 trip)

---

## 常見任務速查

| 任務 | 起手檔 |
|---|---|
| 改規則匹配優先 | `shared/rules.ts:matchEmailWithIndex` + `TYPE_PRIORITY` |
| 改規則學習邏輯 | `background/execute.ts:chooseLearningSignal` |
| 改規則 confidence 上限 | `shared/rules.ts:confidenceCapForType` |
| 加新 rule type | `shared/types.ts:RuleType` + `normalizeSignal` + `applyConfidenceCap` cap + classifier 教 AI 認 + UI 形狀(`RuleForm` / `formatCompoundSignal`) |
| 改同步行為 | `background/sync-engine.ts` + `background/handlers/sync.ts` |
| 加新 settings 欄位 | `types.ts:Settings` + `DEFAULT_SETTINGS` + `storage.ts:sanitizeSettings` + service-worker `setSettings` whitelist + UI;若 per-device,加進 `PER_DEVICE_SETTINGS_FIELDS` |
| 改 AI prompt | `shared/classifier.ts:buildSystemPrompt`(記得 `cache_control: { ttl: '1h' }`) |
| 改近日活動 | `storage.ts:writeFolderActivity`(已有 diff-before-write)+ `folder-activity-filter.ts` |
| 加 popup screen | `popup/App.tsx` state machine + `components/` |
| 加 Options card | `options/App.tsx` — function `XxxCard` + render in App() |
| 加 message handler | `service-worker.ts` 加 case;或 extract 到 `handlers/*.ts` |
| 加常數 | `shared/constants.ts`(不要散在各檔) |
| 處理錯誤 | `logError('source:subtype', msg, context)`(不要再 console.warn-only) |
| 規則庫 UI 改動 | `options/App.tsx` 內的 `RuleLibrarySummaryCard` / `RuleLibraryView` / `RuleAllView` / `RuleRow` / `RuleGroupedTable` / `RuleDetailDrawer` |
