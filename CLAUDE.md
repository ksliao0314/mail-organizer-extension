# Mail Organizer — Claude Code 專案說明

Chrome MV3 擴充工具,把 Outlook 收件夾用「rule + Claude AI」兩段式歸類。權威使用者文件:`README.md`。本檔是給 Claude Code(及其他 AI 助手)接手時看的 architecture / convention / gotcha 速查。

最後一次大改:**2026-11-06**(v1.0.9 單批上限 500 + 分類 / 搬信平行加速;v1.0.10 再審查約 50 項全面修補。更早:2026-11-05 的 v1.0.7 急救批 + v1.0.8 P1/P2 全面修補。逐項見 `CHANGELOG.md`)。本檔下方標 (2026-05-xx) 的段落若與「2026-11-05 / 2026-11-06 更新」衝突,以後者為準。

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

代碼大小快查(2026-11-06):`service-worker.ts` ≈ 4700 行、`options/App.tsx` ≈ 7800 行、`popup/App.tsx` ≈ 4200 行、`sync-engine.ts` ≈ 3200 行、`shared/rules.ts` ≈ 3900 行、`background/execute.ts` ≈ 4100 行、`shared/classifier.ts` ≈ 2200 行。1504 tests across 111 test files。

---

## 分類管線(5 層 fallback)

每封信走以下順序、先命中先勝:

1. **Thread Memory** — `service-worker.classifyPreflight` Step 0。同 ConversationId 或同 normalized subject 之前歸過、直接套用(信心 0.95 / 0.85)
2. **Rule Index** — `rules.matchEmailWithIndex`。**2026-11-05 起為編譯好的 tier 順序**(不再單純依 type),先命中先 return:
   1. `case_code`(Latin 案件代號) → 2. `compound`(最長主旨條件優先) → 3. 法院字號型 subject 規則
   4. **user_manual**(domain → subject_keyword → sender;手寫規則不再輸給自動 domain 規則)
   5. 自動 `domain`(任一地址;**本機信箱網域的地址不觸發自動 domain 規則**,`buildRuleIndex(rules, { ownDomains })`,網域來自 GET /me 快取於 `mailboxIdentity`)
   6. 內文案號 body pass(AI 路徑抓完整內文後) → 7. 自動 `subject_keyword`(longest-first) → 8. 自動 `sender`(只看 From)
3. **Claude AI** — `classifier.classifyBatch`。對 unmatched 跑 LLM
4. **Threshold gate** — AI 信心 < `aiConfidenceThreshold`(預設 0.5) → 標 `source: 'unresolved'`、不執行。AI 呼叫失敗 / 漏回的占位項另標 `unclassified: true`,執行後**不**寫入 skipHistory(除非使用者明確選保留)
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
| `folderCache` | Outlook 資料夾樹快取;`nonFiling`(2026-11-06)= 不能歸檔的 well-known 資料夾 id(草稿 / 寄件匣 / 已刪除的郵件的子資料夾),取代 v1.0.8 的 `excludedFolderIds` | — |
| `folderActivity` | 近日活動 ledger | 200 |
| `folderActivityRefreshAt` | popup auto-refresh 節流 | — |
| `conversationMemory` | ConvId → folder map | 5000 |
| `subjectMemory` | normalized subject → folder + conflictCount | 3000 |
| `metrics` | 累積統計 | — |
| `undoSnapshot` | 30 秒撤回快照 | 過期清除 |
| `weeklyDigest` | 週報 snapshot | — |
| `syncBackups` | 本機 snapshot:只在 replace 合併 / 遠端 wipe / 首次啟用 / 還原 / 撕裂修復前寫;固定 slot 每種 trigger 一份(「套用他機全部刪除前」、「升級同步格式前」);upgrade / first-enable 含墓碑(2026-11-06) | 5 + pinned |
| `syncLastError` | 上次同步失敗紀錄 | next success 清空 |
| `remoteWipeNotice` | 收到另一台機器 wipeMarker 後的本地通知 | 用戶 dismiss 清除 |
| `errorLog` | 中央 SW 錯誤集中 | 200 entries |
| `activeDays` | 本機有執行批次的 UTC 日('YYYY-MM-DD'),stale sweep 的時鐘(2026-11-06 起只記本機;其他機器的在 `fleetActiveDays`) | 400 |
| `lastSweepAt` | 上次 stale sweep 時間(ms),啟動補跑的 20h 閘門 | — |
| `mailboxIdentity` | 本機信箱 `{ address, domain, at, accountId? }`(GET /me,快取 7 天;own-domain 排除用)。2026-11-06 起依 MSAL 帳號(homeAccountId)綁定,換帳號就重查;學習器與初始掃描一律經 `src/background/mailbox-identity.ts` 的 `getAccountCheckedIdentity()` 讀取 | — |
| `movedByUsLedger` | 本工具搬過的信(目的地新 id;harvester 排除自己搬的);**只在 harvestEnabled 時寫** | 5000 + 90 天 |
| `lastHarvestAt` | harvester 水位(掃描**開始**時間 − 10 分) | — |
| `syncLocalState` | 同步本機狀態:上次合併的 pushId、已見過的 wipe id、上次同步的 settings 快照(三方合併用);2026-11-06 加 `v4Adopted` / `legacyLastSyncAt`(升級採用)、`foreignSettings`(別台的未知模型)、`counterPushDueAt`(計數合併推送)。永不同步 / 備份 / 匯出 | — |
| `syncTornStreak` | 雲端撕裂讀取連續次數(依 meta pushId);≥3 次且跨 ≥60 秒才允許修復推送(2026-11-06) | — |
| `refusedEmails` | Claude 拒絕分類過的信 `{ [emailId]: { at, category, model } }`,後續批次不再送 Claude、也不佔批次名額(換模型會重試)。本機限定,不同步 / 備份 / 匯出(2026-11-06) | 500 + 7 天 |
| `refusalSuspicion` | 整批共用內容被拒的判定 `{ model, promptKey, category, at }`:同模型 + 同 prompt 前綴在期限內不再逐封拆分,任何成功回答即清除。本機限定(2026-11-06) | 6 小時 |
| `fleetActiveDays` | 從同步收到的其他機器使用日;`getActiveDays()` = 本機 ∪ fleet(同步規則用),`getLocalActiveDays()` 只含本機(auto_scan 規則用)(2026-11-06) | 400 |
| `ruleCounterBaseline` | 升級後第一次清理記下的各規則命中 / 改歸計數基準(含 epochs);high-error-rate 要基準後 ≥20 次觸發才判斷。只有 `statsResetAt` 變新才算重設,其他下降重新定基準;同步 / 還原 / 匯入帶進來的計數由 `mutateRules` 的 `coverArrivingCounters` 先納入基準。本機限定(2026-11-06) | — |

### `chrome.storage.session`(SW 生命週期)

`cachedToken`(含 `accountId`) / `lastTokenError`(token.ts,getStatus 轉給 popup 顯示真正原因) / `aiClassifyProgress`(帶 `runId`) / `classifyRun`(目前分類 run 的擁有者) / `popupState` / `preflightCache` / `classifyStage` / `executeState` / `executeFinalize`(批次收尾步驟 marker,SW 中途被砍可從斷點重放;2026-11-06 加 `intents` / `statAt`(計數 exactly-once)與 `subjectPrior`(主旨記憶重放)、`learningIntent` / `ruleDisables`(學習步驟與 one-strike 停用 exactly-once,撤回用)) / `executeInFlight`(搬移 / 刪除 / 建資料夾送出前寫入的進行中索引;SW 被砍時這些列標「可能已完成」) / `mailboxIdentityFailure`(/me 查詢失敗冷卻,`{ byAccount }`) / `scanState` / `folderTreeStale`(含 `logged`)

### `chrome.storage.sync`(跨機器,100 KB quota)

| Key | 內容 |
|---|---|
| `syncMeta` | schemaVersion(**4**)/ sourceMachineId / updatedAt / chunkCounts / `pushId` / `contentHash` / `recentPushes[]` / `wipeMarker?`(含隨機 `id`) |
| `syncSettings4` | settings minus per-device fields(帶 pushId) |
| `syncRows4_0..N` | 規則 compact rows(user_manual + ai_confirmed + ai_overridden;預算放不下時,**最近 30 天內停用**的規則以 stub 形式保證送出(2026-11-06 前是所有停用規則);可帶 `sr`(statsResetAt)) |
| `syncFolders4_0..F` | 資料夾字典(hash key) |
| `syncTombs4_0..M` | 墓碑(最新優先;一般最多佔預算 15%,近 30 天的墓碑可擠掉低優先規則列、最多到 40%,2026-11-06) |
| `syncActivity4_0..K` | 最近 20 個 folder activity |

v1–v3 的舊 key(`syncSettings` / `syncRules_*` / `syncRulesC_*` / `syncFolders_*` / `syncTombstones_*` / `syncFolderActivity_*`)只讀不寫,v4 第一次 push 時清掉。每個 v4 chunk 帶 meta 的 `pushId`,讀到不一致(撕裂讀取)就整次放棄、退避重試。

---

## 跨機器同步(sync-engine.ts)

### 2026-11-05 v4 重點(與下方舊段落衝突時以此為準)

- **逐條規則合併**:`Rule.editedAt`(ms)只由 `mutateRules` 在語意欄位(type / signal / target / confidence / enabled / source / autoDisabled* / orphaned)改變時蓋章;命中計數不蓋章。合併:語意欄位取 editedAt 較新者(平手取雲端)、matchCount / overrideCount / lastUsedAt 取最大、createdAt 取最小。同步 pull / restore 用 `mutateRules(fn, { preserveEditedAt: true })`。
- **單一佇列**:push / pull / restore / clearCloud 共用一個 queue;push 前若雲端 pushId 不是本機上次合併過的就先 pull(union);pull 期間的本機變更延後推送、不再丟棄;settings 用上次同步快照做三方合併。
- **墓碑看時間**:墓碑只壓 `max(createdAt, editedAt) <= deletedAt` 的規則版本(雙向);刪除時用 `tombstoneDeletedAt()` 蓋 `deletedAt = max(now, editedAt+1, createdAt+1)`,時鐘偏差也能傳遞刪除。墓碑 key 一律經 `tombstoneKey()` 重新正規化(相容舊格式)。
- **停用一定傳得出去**:預算放不下的停用規則以 stub `[id, editedAt, reason, at]` 送出、不被截掉;stub 不會在對方新建規則。(2026-11-06 起只保證最近 30 天內編輯過的停用,見下方 v1.0.10 補充)
- **手動「立即下載」永遠 union**(v1.0.7)。
- **wipeMarker 有隨機 id**:每台機器記錄已見過的 id、只套用沒見過的;首次參與的機器只記錄不套用;停用同步會重置參與狀態。
- **錯誤分類**:`MAX_WRITE_OPERATIONS_*` 是速率限制(退避後重試),不是容量問題;push 失敗會依指數退避自動重試。
- 混版:v1.0.8(v4)可讀 v1–v4 雲端並升級成 v4;v1.0.7(v3)遇到 v4 雲端會拒絕並提示升級。

### 2026-11-06 v1.0.10 補充(與上面衝突時以此為準)

- **升級採用(adoption)**:每台機器第一次 v4 合併前,先建 pinned「升級同步格式前」備份(規則 + 墓碑 + 設定)。本機活著的規則若只被「自己的墓碑」或「v3 已處理過的墓碑(deletedAt ≤ 舊 lastSyncAt)」壓住 → 保留,editedAt 蓋到 deletedAt+1 之上讓所有機器收斂;user_manual 遇到「只在新正規化後才相撞」的墓碑也保留。
- **`orphaned` 是本機狀態**:不在語意欄位、不蓋 editedAt;本機孤兒版本不會在合併中勝出;合併採用不同目標時清掉孤兒旗標讓 reconcile 重判。
- **統計 epoch**:`Rule.statsResetAt`(row key `sr`),復活重設計數時蓋章;同步合併與 `dedupeRulesByKey` 的 matchCount / overrideCount 只取最新 epoch,同 epoch 取最大。
- **還原 = 新的編輯**:與目前不同(或本機沒有)的還原版本蓋新 editedAt(高於對應墓碑);備份之後才出現的同步規則以墓碑移除;回傳 `removedCount`。
- **撕裂雲端修復**:同一 meta pushId 連續 ≥3 次撕裂讀取且跨 ≥60 秒 → 先建「修復雲端資料前」備份、再以本機重寫雲端(帶著本機沒套用過的 wipe 時拒絕);meta 擁有者看到外來 chunk 會主動修復。
- **預算**:停用 stub 只保證最近 30 天編輯過的停用;近 30 天的墓碑最多可佔 40% 預算,只擠掉低優先列(從未命中的 ai_confirmed、停用規則)。
- **計數合併推送**:只有命中數變動(`meta.semanticHash` 不變)時最多每 20 分鐘推一次(alarm `mail-organizer:sync-counter-push`);語意變更、墓碑、設定、近日活動、新使用日、`pushNow` 立即推。`contentHash` 公式不變,與 v1.0.8 的 echo 偵測相容。
- **使用日同步**:`activeDays` 以 settings item 的 `d` 欄位(差分編碼)同步,pull 時 union(`mergeActiveDaysFromSync`),各機器 stale sweep 的時鐘一致。
- **參與狀態**:啟用同步時 `beginSyncParticipation()`;沒有 v4 參與狀態的機器把現有 wipe id 記為已見、不套用(舊 lastSyncAt 只當升級採用的分界)。
- **未知模型**:別台選了本機不認得的 `claudeModel` → 記在 `foreignSettings`、推送時原樣送回,本機模型不變。
- **mutateRules 先讀墓碑**:`stampEditedAt` 對新規則與重新啟用(false→true)蓋到該三元組最新刪除時間之上(時鐘偏差不再吞掉重建的規則);重複 id 先依物件身分、再依出現順序配對。
- **墓碑依寫法(第三輪)**:每筆墓碑由 `ruleTombstone()` 建立並帶 `spellingHash`(被刪規則的 legacy 寫法指紋,線上為第 5 個欄位);storage 以 (三元組, spellingHash) 為鍵(`mergeTombstoneSlots`),較新的 deletedAt 勝。帶 spellingHash 的墓碑只壓同寫法的規則(穩態合併與升級採用皆然),沒有 spellingHash 的(v1.0.8 / v1.0.7 / 匯入)壓所有寫法;v1.0.8 回推的去寫法副本會併回原本有寫法的那筆。第二輪的「刪除時重蓋存活雙胞胎」已移除。v1.0.8 機器不認寫法,在它那邊刪一種仍會刪掉另一種,但不會擴散、也不會來回同步。
- 向下相容:`sr` / `d` / `semanticHash` / `daysHash` / 墓碑第 5 欄全是附加欄位,v1.0.8(v4)會忽略;pushId 撕裂檢查不變。

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

`SYNC_SCHEMA_VERSION = 4`(2026-11-05)。v3(2026-07-15)= compact rows + 資料夾字典;v4 = editedAt / pushId / 停用 stub / 精簡 compound 編碼 / 新 reason code(`override` 4、`superseded` 5)。**reason / type code 只能 append**,並且要跟第一個寫入它的版本一起出貨;未知 type code 的 row 會被略過,未知 reason 只丟 reason。

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
| **自動沉寂** | `stale-sweep.ts:runSweep` → `autoDisableStaleRules`(每日 alarm;SW 啟動時若距上次 ≥ 20h 補跑一次;Options「立即執行」= `force`) | 錯誤率以**所有觸發**計(`ruleFires = matchCount + overrideCount`、`ruleErrorRate`;2026-11-05 前只算被接受的命中,越常錯的規則反而活越久)→ `enabled=false` + `autoDisabledReason: 'high-error-rate'` + `autoDisabledAt` |
| **stale 直接刪** | 同上 sweep | `matchCount === 0` 且 createdAt 之後已過 100 個**使用日**(full_subject 類 45) OR `matchCount > 0` 且 lastUsedAt 之後已過 100 個**使用日** → **完全刪除**(不留 tombstone、不寫 audit、不 console)。**使用日** = 有執行批次的 UTC 日:本機 `activeDays` ledger ∪ 從同步規則的 lastUsedAt / AI 規則 createdAt 推導(`deriveActiveDaysFromRules`)。閒置期間不計 — 2026-11-05 前用日曆天,停用 ~100 天回來整個學習庫被刪光、還原備份後下次喚醒又被刪。設計理由:長期沒命中等於不存在,留痕跡只汙染信任 |
| **假案號清理** | 同上 sweep(`pseudoCasePurgeTransform`,2026-11-06) | 案號訊號是「1XX度N」(年度編號被舊解析誤當案號)的自動規則 → 完全刪除(不留墓碑);user_manual 不動。其他「年度+非字別」的舊訊號(如 113財報1)無法證明是假的,保留到 stale sweep |
| **legacy token 清理** | 同上 sweep | 2026-05-27 之前用 tokenize 路徑生的 subject_keyword 規則(`isLegacyTokenSubjectRule`):createdAt < `REDESIGN_CUTOFF_ISO` 的 ai_confirmed / ai_overridden subject_keyword 規則 → 設 `autoDisabledReason: 'legacy_token'` |
| **自動復活** | `addRulesFilteringTombstones` dedup(兩段式,2026-11-06)+ learners 的 `isRevivableAutoDisabled` | 同 type+signal+target 對應到 `override` / `superseded` / `high-error-rate` 自動停用的規則 → 復職(high-error-rate 重新起算準確率並蓋 `statsResetAt`);**例外**:同一錨點(網域 / 寄件人,含錨定在它上的 compound)還有活規則指向別處、或同批候選 / 復活互相衝突時不復活;使用者停用、`case_closed`、`legacy_token` 仍擋住重學 |
| **被改歸 / 被取代** | execute(`override`:案號類規則命中的信被你**改歸到別的資料夾**,一次即停用;刪除 / 保留不算)、learning(`superseded`:廣規則被更精準規則取代) | 帶 `autoDisabledReason` + `autoDisabledAt`;停用滿 60 個使用日由 sweep 硬刪 |
| **使用者手動 toggle** | UI 按鈕 / 表單 / 批次(一律 `withToggle`) | 停用清掉 auto-disable 痕跡(所以不會被學習復活、也不會被休眠清理刪掉);重新啟用 `case_closed` 規則時保留 case_closed 當豁免標記、停用豁免規則則清掉 |
| **錯誤率基準** | `ruleCounterBaseline`(第一次清理建立,本機) | high-error-rate 需要基準之後 ≥20 次觸發、至少一半被改歸(舊版把「刪除」也算改歸);計數低於基準視為重設(復活),以新計數判斷 |
| **撤回** | `executeUndo`(H2 原則:撤回 = 這批的歸檔是錯的) | 有成功搬回的歸檔時,還原這批學習對規則庫的所有改動:兩個學習器新建的規則刪除(`mintedRuleIds`);復活的規則回到休眠(`revivedRules`);被 superseded / override 降級、或 one-strike `override` 停用的規則恢復(`disabledRules`)。受批次後指紋保護(之後手動改過的不動),計數不動 |
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
- **法院字號全形必須有「字第」(2026-11-06)** — `COURT_CASE_FULL_RE` 要求「字」緊接「第」(可有空白),「113年度財報第1季」「113年度第1次董事會」不再解析;`DATE_STOPWORD_RE` 含「度」。簡寫(112訴204)不設字元黑名單:財 / 季 / 午 / 報都出現在真實字別(家財訴、司執季、午字、報字)
- **Outlook DisplayName 含 `/` 編碼為 U+FF0F** — `encodeFolderName` / path 比對都用 fullwidth solidus
- **OData filter PascalCase** — `From/EmailAddress/Address`、不是小寫;`autoUpgradeConflictRules` 若 server filter 失敗會 fall back 到 client-side filter
- **wipeAllRules abort on cloud-clear failure** — 雲端清除失敗時不動本機,避免 cloud 滿、local 空的不一致狀態
- **REDESIGN_CUTOFF_ISO = '2026-05-27T00:00:00.000Z'** — `isLegacyTokenSubjectRule` 用這個分界線辨識舊 tokenize 路徑生的規則
- **「刪除」= 移到「已刪除的郵件」(2026-11-05)** — 以前 `DELETE /me/messages` 會進 Recoverable Items(使用者在「已刪除的郵件」找不到)。現在 POST /move 到 well-known `DeletedItems`(400 時改用解析出的資料夾 id),並且進撤回快照(`deleted: true`)
- **Outlook 並發上限 4** — `OutlookApi.request()` 內建 FIFO semaphore(Exchange 每信箱並發限制);`TokenBucket` 改為預約制(v1.0.9 起 8/s、burst 40,`OUTLOOK_REQUESTS_PER_SEC` / `OUTLOOK_REQUEST_BURST`)。不要再在呼叫端自己加 limiter
- **AI 分組平行(v1.0.9)** — `classifyAi` 最多 `AI_PARALLEL_CHUNKS`(3)組同時呼叫 Claude。第一組先送、等它的回應**開始串流**(`classifyBatch` 的 `onResponseStart`,prompt cache 從此可讀)才放行其餘組;結果依組別順序併入 aiPlan,`completedChunks` 只算連續已併入的前綴;任一 429/529 之後整次退回一次一組;run-fatal 錯誤停止排程、已完成的組(含失敗組之後的)保留、其餘標 `unclassified`。每組的處理(`processChunk`)只回傳結果、不丟例外
- **搬信平行(v1.0.9)** — `runClaimedExecute` 先逐筆跑 new_folder(後面的 move 可能指向它建的資料夾),其餘最多 `EXECUTE_PARALLEL_ITEMS`(3)筆同時;取消在每個 worker 的下一筆邊界生效,沒開始的列標 cancelled
- **收件匣掃描預算(v1.0.9)** — `inboxScanBudget(batchSize, ghosts)`:每頁 100–200、總掃描 = 3× 批次(500–2000);`scanInboxForBatch` 跨頁以 Id 去重($skip 分頁在新信進來時會重列上一頁尾端)。批次上限 `BATCH_SIZE_MAX = 500`
- **特殊資料夾(2026-11-06 第二輪定案)** — 資料夾樹不含「已刪除的郵件」(根與子樹)、「草稿」「寄件匣」,也不列舉它們的子樹;「垃圾郵件」「寄件備份」(含子資料夾)在樹裡、根節點標 `wellKnown`。五個 well-known 根的 id 記在 `folderCache.wellKnown { roots, complete }`;查詢不完整時快取只新鮮 60 秒、reconcile 以 `treeComplete:false` 執行(不改任何規則)。v1.0.8(`excludedFolderIds`)與第一輪(`nonFiling`)格式的快取首次使用時重抓。
  - reconcile:目標是「已刪除的郵件」根 → 保留(使用者的自動丟垃圾桶規則);其他不在完整樹裡的目標 → 孤兒(含在 OWA 被刪、移進垃圾桶的資料夾)
  - execute:目標 = 垃圾桶根 → 以 `moveToDeletedItems` 執行、狀態 `deleted`(不學習、不記憶、不進近日活動,規則命中照計);垃圾桶 / 草稿 / 寄件匣底下 → 拒絕並給明確訊息;資料夾快取超過 10 分鐘(`FOLDER_TREE_VERIFY_AFTER_MS`)時 DestinationGuard 會即時走 parent chain 確認樹內目標(每批 memo、平行 worker 共用),拒絕時以 `invalidateFolderCache()` 蓋 `folderCache.invalidatedAt`,下次 preflight 重抓
  - AI:`buildFolderBlock` / 近期資料夾提示 / 範例 / 案號對照只列 AI 可見資料夾(`flattenAiVisibleFolders`,排除垃圾郵件 / 寄件備份子樹與樹外路徑);AI 指到清單外或排除清單的路徑 → `skip` + `unresolved`(保留 AI 原始選擇)
- **Outlook 請求 60 秒逾時(2026-11-06)** — `OUTLOOK_REQUEST_TIMEOUT_MS`,經 `linkAbortSignals` 與呼叫端的 signal 連動;GET 逾時照網路錯誤重試,非冪等請求逾時或 2xx 本文卡住 → `uncertain`;呼叫端取消維持原語意;slot 一定釋放
- **權杖依資源排序(2026-11-06)** — `msal-token-reader`:outlook.office.com / office365 / outlook.cloud.microsoft(可帶 :443)優先,其次其他非 Graph 的 Mail.ReadWrite;Graph 與 bare scope 永不採用;NO_TOKEN 診斷列出看到的資源(不含權杖)。函式經 `chrome.scripting` func 注入,**必須自給自足**(不能 import、不能用外層變數)
- **keep-alive 心跳(2026-11-06)** — 只要有 hold,每 20 秒呼叫 `chrome.runtime.getPlatformInfo` 重設 idle 計時;最後一個 release 清掉,alarm 保留當備援
- **Claude 拒答隔離(2026-11-06)** — REFUSAL 時在該組的 lane 內二分找出被拒的信(每組最多 12 次、每次 run 最多 36 次額外呼叫);被拒的信記進 `refusedEmails`(依模型、7 天),錯誤紀錄不寫主旨。第二輪:每個 run 有一個 refusal gate(`createRefusalGate`)— run 還沒有任何成功回答前,先送一封合成測試信(`sharedPromptProbeEmail`)驗證共用 prompt,其他 lane 等結果;測試信也被拒 = 共用內容造成,不拆分、記 `refusalSuspicion` 6 小時;warm gate 改在第一段答案文字才打開
- **串流 90 秒閒置看門狗(2026-11-06)** — 連 header 都算;90 秒沒有任何位元組 → `NETWORK`(可重試);使用者取消仍是 `ABORTED`
- **重試分種類(2026-11-06)** — `withClaudeRetry` 的 backoff(429/529)與 transient(網路 / 5xx / 串流)各有自己的排程與預算
- **費用逐次計價(2026-11-06)** — 讀取 serving `model`、`usage.iterations` 與 fallback 區塊;`billedUsage()` 依實際執行的模型計價,`ClassifierUsage.costUsd` 有值時 popup 直接用
- **5.x prompt(2026-11-06)** — 5.x 的 reason 是決定後的一句說明(schema 把 reason 放最後,避開 reasoning_extraction 拒答);4.x 維持舊的 reason-first prompt / schema(byte-for-byte)
- **內嵌面板配對(2026-11-06)** — owa-fab 向 SW 要一次性 nonce(`issuePanelPairingNonce`,content script 白名單)放進 `parent-hello`;面板 iframe 用 `redeemPanelPairingNonce` 確認是同一分頁發的才接受 port;配對前不送任何資料夾路徑。nonce 只在 SW 記憶體(60 秒、每分頁 16、總共 512)。第二輪起 popup 與 options(owa-fab 的設定 modal 也內嵌它)都由 `src/components/ui/bridge-gate.tsx` 包住:被內嵌時要等配對驗證完成才渲染(防誘導點擊),10 秒仍未配對就提示重新整理 Outlook 分頁
- **計畫清單鍵盤(2026-11-06)** — 清單內(`data-plan-list` / `data-plan-row-index`)有游標時 Space = 勾選、Enter = 展開 / 收起游標列;清單外的按鈕(執行)保留原生 Enter;「待確認」chip 與「依資料夾」切換滑鼠點擊不取焦點。第二輪:游標以 emailId 追蹤(排序 / 合併 / 群組拆開不會換到別封);`createFocusOriginTracker` 讓 Tab 移入的列成為游標,Tab 到的非列按鈕(全選此組等)原生觸發;Esc 先清游標、再按才收面板
- **非冪等 retry 是禁忌** — `outlook-api.ts:request()` 對 POST / DELETE / PATCH **不** retry 網路錯誤 + 5xx,只 retry 429。若硬把這條規則破掉,POST `/move` 第一次成功但 response 掉了的場景會在重試時 404、UI 顯示成失敗、使用者誤以為訊息沒搬走。`OutlookError.uncertain = true` 是給這類「可能成功」case 用的 marker
- **404 ErrorItemNotFound on move/delete = soft skip,不是 error** — `executeItem` 用 `isAlreadyMovedError(e)` helper 偵測、回傳 `status: 'skipped'`。原因:訊息已不在那個 Id(可能我們上一次 POST 其實成功了、或被其他規則 / 帳號處理過)。重試只會再 404
- **(2026-11-05 起改用 run id)** 分類 progress 現在有擁有者:preflight 發 `runId`、classifyAi 只服務目前的 run、迴圈每次呼叫 Claude / 寫 progress 前都檢查;取消 = `cancelClassify`。部分執行時 SW 記下已執行的 email id,progress 寫入與讀取都排除它們;消費 progress 依 runId 而非下方的 startedAt 啟發式。下方原文保留作歷史脈絡:
- **「已歸檔的信混進下一批」的真正主因是 progress 殘留,不是 Outlook 延遲** — AI classify progress 躺在 session storage,執行完不會被清(confirmExecute 只清 popupState)。`continueToNextBatch` 若不檢查就消費,等於把剛執行那批的 plan 原封復活——該路徑不抓收件匣,任何 fetch 端過濾都救不了。守門:**只消費 `progress.startedAt > executeState.startedAt` 的 progress**(執行按鈕 gated on aiPending ⇒ 殘留恆舊、prefetch 恆新,判別無漏洞);SW `clearExecuteState` 另把可證明過期的 progress + preflight cache 一併清掉(堵 reopen-resume 入口)。classifyAi 的 `startedAt` 必須**不可變**(迴圈前取一次),否則邊界模糊
- **Outlook 收件匣列表是最終一致 → 用 recentlyProcessed 帳本過濾(次要防線)** — move/delete 後幾秒〜一兩分鐘內,`listInboxMessages` 仍可能用舊 id 列回剛搬走的信。`execute.ts` 批次結束記下 moved/deleted/folder_created **以及 404 already-moved(`ItemOutcome.alreadyGone`)** 的 emailId 進 `recentlyProcessed`(TTL 15 分、`storage.ts`),`classifyPreflight` + `peekNextBatch` 先讀帳本、**over-fetch `top = batchSize + min(幽靈數,100)`** 再過濾、最後裁回 batchSize。**uncertain 軟跳過的信不記帳本**(讓真沒搬動的能重試;已確認死亡的 404 則要記,否則延遲拖長會無限輪迴)

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

`tests/` 目錄、vitest + jsdom + mocked chrome.storage。**1504 tests across 111 files**(2026-11-06,v1.0.10)。共用 helper:`tests/helpers/sw-harness.ts`(載入真正的 SW dispatcher,用 UI 實際訊息格式測 contract)、`classify-env.ts` + `claude-sse.ts`(classify 端到端 + SSE fixture)、`sync-sim.ts`(多機器同步模擬)、`ui-harness.ts`(jsdom 渲染真正的 popup / options)、`active-days.ts`。重要測試:

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
└── tests/                          ← 1504 tests (2026-11-06)
```

2026-11-06 新增:`src/background/refusal-isolation.ts`(拒答二分隔離)、`src/shared/refusal-ledger.ts`(`refusedEmails`)、`src/background/handlers/panel-pairing.ts`(面板配對 nonce)、`src/popup/crash-recovery.ts`(錯誤畫面「捨棄暫存計畫並重新載入」)、`src/options/rule-accuracy.ts`(全頁共用的準確率計算)。

2026-11-05 新增:`src/shared/claude-models.ts`(模型表:id / 世代 / 價格 / 是否支援 structured outputs、server-side fallback)、`src/shared/msal-token-reader.ts`(MSAL token 讀取,content script 與測試共用)、`src/background/handlers/rules-bulk.ts`(`bulkRuleAction` 與批次規則操作,一次 mutateRules)、`src/popup/keyboard.ts`、`src/popup/web-storage.ts`(包好 try/catch 的 local/sessionStorage)、`src/options/rule-list.ts`(refresh 時保留未變動規則的物件身分)、`src/components/ui/error-boundary.tsx`、`src/fonts.css`(只打包 Latin 字型子集)。

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
9. **wipeMarker 寫入時不抹除舊 marker** — `doPush` 在 normal push 時保留 `cloudMeta.wipeMarker`、只有 wipe push 才覆蓋(讓第三台機器後加入也能 trip)。2026-11-05 起每個 marker 有隨機 id,判斷「是否已套用」看 id、不再比較兩台機器的時鐘
10. **content script 永不碰 chrome.storage**(2026-11-05) — `settings` 含 `claudeApiKey`;owa-fab 曾 `console.log` 整個 settings、且其 `storage.onChanged` 每次 settings 寫入都收到含 key 的新舊值。現在 owa-fab 只透過 `getOwaFabVisibility` 拿布林值、由 SW 推 `owaFabVisibility`。`tests/ui-sw-contract.test.ts` 掃 `src/content/*.ts` 守門
11. **手動拉取(立即下載)永遠 union**(2026-11-05) — `doPush` 的位元組預算會刻意把低優先規則留在雲端外(只存本機),replace 模式會把這些唯一副本刪掉。`replace` 只留給明確需要「雲端精確狀態」的內部路徑;刪除靠 tombstone 傳遞
12. **stale sweep 用使用日、不用日曆天** — 見「規則生命週期」;新增任何會「依時間刪規則」的路徑都要問:使用者沒在用時這個時鐘該不該走
13. **`editedAt` 只由 `mutateRules` 蓋章**(2026-11-05) — 其他程式碼不要自己設;sync pull / restore 用 `preserveEditedAt`。2026-11-06 起的合法例外都在 mutateRules 交易內:`stampEditedAt` 把新規則 / 重新啟用蓋到墓碑之上;同步的升級採用與還原以明確選項設定 editedAt
14. **SW 只服務擴充頁面** — dispatcher 檢查 sender:`sender.url` 不在 `chrome-extension://<id>/` 底下的(content script)只能用 `CONTENT_SCRIPT_MESSAGE_TYPES` 白名單(目前有 `getOwaFabVisibility` 與 `issuePanelPairingNonce`)。新增 content script 訊息要加進白名單
15. **新增 message type 要加進 `SW_MESSAGE_TYPES`**(`src/shared/messages.ts`) — `handle()` 的 switch 是 exhaustive、UI 的 `send()` 也用這個型別,漏加會編譯失敗
16. **墓碑一律 `tombstoneKey()` / `tombstoneDeletedAt()`** — 前者相容舊格式 signalNorm,後者避免時鐘偏差讓刪除傳不出去;2026-11-06 起一律由 `ruleTombstone()` 建立(帶 spellingHash)
17. **`orphaned` 是本機狀態**(2026-11-06) — 不蓋 editedAt、不讓本機孤兒版本在同步合併中勝出
18. **計數只在同一個 `statsResetAt` epoch 內合併**(2026-11-06) — 同步、去重都一樣;復活重設的計數不能被舊副本的計數蓋回

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
