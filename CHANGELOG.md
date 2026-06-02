# Changelog

時間軸最新在上。每次「規則 / 同步 / Options 結構 / 安全」相關的重大變更都記在這。實作細節請看 `CLAUDE.md`、使用者文件看 `README.md`。

---

## 2026-06-03 — 全面 bug 掃描(5 agent 平行 + 自驗)13 項修補

並行 5 個 agent 掃過 sync / execute+api / rules / SW+storage / UI,自驗後確認 13 真 1 誤報(`recoverStaleExecuteState` 被指可與 live executor 並發 → 駁回,因 onInstalled/onStartup/module-load 都在全新 SW 跑、舊 executor 已隨更新 tear down)。

### P1

- **F1 跨機器 wipe 不再誤殺新機器**:`doPull` 的 wipeMarker 偵測加 `settings.lastSyncAt !== ''` 守衛。首次啟用同步的機器(lastSyncAt 為空)其 union pull 本意是保留本地規則,卻會踩到 cloud 殘留的舊 wipeMarker、把整個規則庫清空。現在從未同步過的機器豁免於 wipe — marker 只對「已在同步集裡」的機器生效。
- **F2 wipe 套用後立即推進 lastSyncAt**:遠端 wipe 的破壞性 mutation commit 後、馬上 `setSettings({ lastSyncAt: marker.at })`,在後續 chunk 讀取 / merge(可能拋網路 / quota 錯)之前。否則一旦下游拋錯,wipe 已套用但 lastSyncAt 未前進 → marker 保持「新鮮」→ 下次任何 remote push 都重複觸發、反覆清掉期間重建的規則。
- **F8 dedupe 不再硬刪 user_manual**:`dedupeRulesByKey` 的 comparator 把 user_manual 排在 enabled 偏好「之前」。舊版 enabled-first 排序會讓「停用的 user_manual」輸給「啟用的 auto 規則」(同 triple)被硬刪 — 無墓碑、無 audit、無痕跡地摧毀使用者刻意建立的規則。其他來源仍維持 enabled-first。

### P2

- **F3 遠端 wipe 不清本機 tombstones**:`doPull` 遠端 wipe 套用時拿掉 `clearAllRuleTombstones()`。tombstones 是本機自己的刪除意圖,使用者沒在這台點「全部刪除」,清掉會讓本機刪過的規則之後被第三台機器推回來復活。本機 `wipeAllRules`(使用者主動)仍清 tombstones。
- **F5 網路中斷的 move 不再顯示紅 error**:`startExecute` catch 偵測 `OutlookError.uncertain`(非冪等操作網路 / 5xx 失敗、伺服器可能已 commit),標 `skipped` + 「網路中斷,這封可能已移動也可能還在收件匣,請確認」誠實訊息,而非紅色 error。完成 2026-06-02 retry 修復的最後一塊 — 杜絕「已成功移動卻叫我再歸檔」殘留症狀。
- **F11 recordRuleEvents 加 mutex**:審計日誌的 read-append-write 無序列化,並發 append(preflight reconcile + Options 刪除)會互相覆蓋丟事件。包進 `withHistoryLock`。
- **F12 wipeAllRules 走 mutateRules**:`全部刪除` 與 dev 用 `clearAllRules` 改用 `mutateRules(()=>({next:[]}))` 取代 raw `setRules([])`,避免並發 batch 的 bumpRuleHits 在 wipe 後復活規則(wipeInFlight 只擋第二次 wipe、不擋其他寫入者)。
- **F15 popup 鍵盤導覽走 displayItems**:`focusedIndex` 改 index `displayItems`(渲染列)而非 `filteredItems`。游標不再落在折疊隱藏的對話 sibling 上、d/s 不再改到看不見的信。
- **F16 drawer focus trap 不搶巢狀 ConfirmDialog**:Tab 處理器偵測到 focus 在 `role="alertdialog"` 內時 bail,讓確認對話框管自己的 focus、鍵盤可 Tab 到刪除確認鈕。
- **F17 ActionsMenu 方向鍵跳過 disabled**:方向鍵 / Home / End 過濾掉 disabled 項,navigation 不再卡在不可聚焦的按鈕上。

### P3

- **F6 Retry-After clamp**:429 / 5xx 的 `Retry-After` 上限 60s(`MAX_RETRY_AFTER_MS`),防惡意大值凍結整批。
- **F9 bumpRuleOverrides 先於 bumpRuleHits**:同批升信心的 gate 讀到本批更新後的 overrideCount,不再用舊值誤升一條使用者正在覆寫的規則。
- **F13 tombstone cap 按 deletedAt**:`setRuleTombstones` 超 cap 時按 `deletedAt` 排序保留最新,而非按陣列位置(re-tombstone 會更新 deletedAt 但保留舊插入位置 → 位置裁切會誤刪較新的)。
- **F18 移除 dead onMouseLeave**:ActionsMenu 空 handler dead code。

### 新增測試

- `tests/sync-wipe-marker.test.ts`:加 F1 first-participant 豁免、F2 lastSyncAt 推進;原「清 tombstone」測試改為「保留 tombstone」(F3)
- `tests/rules.test.ts`:加 dedupe user_manual 神聖性 3 case(F8)

測試總數:**416 tests across 24 files**。

---

## 2026-06-02 — 非冪等 POST retry 修復 + 404-as-skipped

### 根因

`outlook-api.ts:request()` 對所有 HTTP method 都會在網路錯誤 + 5xx 時重試最多 3 次。對 `POST /me/messages/{id}/move` 這類**非冪等**操作:第一次請求其實已經被 Outlook 伺服器收到並完成(訊息已經被移動、Id 已變更),但回應在傳回客戶端的路上掉了(Wi-Fi 切換、TLS reset、暫時性 5xx)。重試用**舊 Id** 再送 → Outlook 回 404 `ErrorItemNotFound` → UI 標成失敗、邀請使用者再重試 → 使用者就看到「明明已經搬走了還顯示要繼續歸檔」的迴圈。

### Fix A — `outlook-api.ts` retry 分流(`isIdempotentMethod`)

- `GET` / `HEAD` / `PUT` / `OPTIONS`(冪等):維持原本網路錯誤 + 5xx 重試 3 次
- `POST` / `DELETE` / `PATCH`(非冪等):
  - **429**(rate limit、伺服器明確「沒處理」)仍重試
  - **網路錯誤** → 立刻拋,`OutlookError.uncertain = true` + 訊息「operation may have completed server-side」
  - **5xx** → 立刻拋,`OutlookError.uncertain = true`
- `OutlookError` 新增 `uncertain: boolean` 欄位讓上層辨識「可能成功、不確定」case

### Fix B — `executeItem` 把 404 ErrorItemNotFound 當 `skipped`

新增 `isAlreadyMovedError(e)` helper(`status === 404 && /ErrorItemNotFound/i`)。在 `executeItem` 的三個移動點(直接 move、move into existing folder、new_folder 後 move)包 try/catch:
- 抓到 already-moved 錯誤 → 回傳 `{ status: 'skipped', message: '訊息已不在原位置(可能已成功移動、或被其他規則 / 帳號處理過)' }`
- 不會被 summary 算成 error、UI 不會邀請使用者再重試
- delete 同樣處理(`api.deleteMessage` 也是非冪等的 DELETE)
- new_folder 後 move 失敗時、保留 `destinationFolderId` / `destinationFolderPath`,讓同批的後續 sibling 仍能 reuse 這個新建的資料夾

### `ItemOutcome` 增加 `message?`

讓 `executeItem` 能在成功 / 軟跳過路徑帶解釋文字到 `ExecuteItemResult.message`,UI 顯示 soft note 而非紅色 error。

### 新增測試

- `tests/outlook-api-retry.test.ts`(9 tests):POST 不在網路錯誤 / 5xx retry、POST 仍在 429 retry、GET 仍在 5xx + 網路錯誤 retry、404 surface 為 `uncertain=false`、`OutlookError.uncertain` 在非冪等失敗時為 true
- `tests/execute-already-moved.test.ts`(7 tests):`isAlreadyMovedError` 對 404+ErrorItemNotFound / 大小寫 / 沒 marker 的 404 / 500 + ErrorItemNotFound / plain Error / 非 Error 值的辨識

測試總數:**411 tests across 24 files**。

---

## 2026-05-27 — 規則庫重新設計 + 學習簡化 + 跨機器 wipe 傳遞 + a11y polish

### 規則學習:先廣後窄

**核心哲學變更**:同一網域的規則,**第一次**遇到該網域時建立 plain-domain 規則(廣);第二次遇到該網域且歸到**不同**資料夾時,自動升級為 compound(網域 + 整段主旨),原 broad 規則自動沉默。

- `chooseLearningSignal` 重寫為 6 個優先級(由窄到廣):
  1. 外部域 + 法院字號 → `compound`
  2. 法院字號(無外部域) → `subject_keyword`
  3. Latin 案件代號 → `case_code`
  4. 可用網域(非 generic provider)+ 衝突偵測 → 預設 `domain`、衝突時 `compound`(網域 + 整段主旨)
  5. 通用 provider(gmail.com 等) + 完整地址 → `sender`
  6. 內部網域 fallback + 主旨可用 → `subject_keyword`(整段主旨)
- 衝突偵測用 `normalizeSignal('domain', r.signal)` 比對(處理 import / manual 帶 `@` 前綴的 legacy 形式)
- `demoteOnly` flag:衝突偵測到但 subject 抽不出可用 signal 時,降為「只 demote 原衝突 broad 規則、不創新規則」

### 主旨即特徵(不再 tokenize)

- 移除舊的 `extractSubjectFeature` / `extractSubjectFeatureLenient` 拆詞路徑
- 改用 `extractSubjectSignal`:把整段主旨正規化(去 Re:/Fwd:、空白合併、轉小寫)當 signal
- 最短長度 `MIN_SUBJECT_SIGNAL_LEN = 3`
- 舊版 tokenize 路徑生成的 subject_keyword 規則(createdAt < `REDESIGN_CUTOFF_ISO = '2026-05-27T00:00:00.000Z'`)會被 `autoDisableStaleRules` 標 `autoDisabledReason: 'legacy_token'`、保留可手動啟用

### Confidence 上限(applyConfidenceCap)

新增寫入時強制 cap、`user_manual` 例外、按 type:
- `domain`: 0.7(廣域、必須低於 compound 才有意義)
- `sender`: 0.75
- `subject_keyword`: 0.9
- `compound` / `case_code`: 0.95

寫入點覆蓋:`upsertRule` / `addRules` / `bumpRuleHits` / `sync-engine.doPull` / `sync-engine.restoreBackup` / `service-worker.importRules`。匯入時若有 cap 觸發,UI 顯示「N 條規則的信心度從 X 降至 Y(符合系統上限)」。

### 規則庫 UI:儀表板 + 全螢幕工作區

Options 頁規則庫從「單一長頁」改為兩段式:

1. **儀表板**(預設):4 卡片(總數 / 啟用 / 衝突 / 休眠) + Top 5 命中規則 + 「打開規則庫」按鈕
2. **全螢幕工作區**(`#rules-library/*`):左側 6 個子分頁 — 全部 / 衝突 / 休眠 / 健康 / 初始掃描 / 歷史

全部規則改為**按資料夾分組**(預設),`viewMode` 可切 grouped / flat、collapsed 狀態存 sessionStorage。每組標題顯示 N 條 / 命中總數 / 平均準確率,hover 露 「全部停用」action。

新元件:`RuleLibrarySummaryCard` / `RuleLibraryView` / `RuleLibraryViewBody` / `RuleAllView` / `RuleRow` / `RuleGroupedTable` / `RuleDetailDrawer` / `RuleLibraryActionsMenu` / `ConfirmDialog` / `ConflictsView` / `DormantRulesView`。移除舊的 `RulesSection`(1742 行)/ `KeywordChipInput` / `RulesBreakdownStrip` / `RuleEffectivenessSection`。

「自動沉寂」全面改名為「**自動休眠**」(降低語意刺耳感)。

### Options 頁 3 tabs(原 5)

- **分類引擎**(原 connection + skip + 內部規則 + AI 設定合併)
- **規則庫**
- **資料 & 同步**(原 sync + import/export + wipe)

Hash redirect:`#connection` → `#engine`,`#skip` → `#data`。`skipFlagged` 設定從 UI 移除、SW 永遠跳過 flagged(這是預設,不需用戶選)。

### 跨機器 wipe 傳遞(wipeMarker)

**問題**:union-mode pull 會保留本機沒在 cloud 出現的同步規則。在 A 機按「全部刪除」後,B 機下一次推送會把它的舊規則送進 cloud,A 機 pull 回來、wipe 等於白做。

**機制**:
- `SyncMeta` 新增 `wipeMarker?: { at: string; byMachineId: string }`
- `wipeAllRules` handler:
  - quiesce + clearCloudState → **失敗時 abort,不動本機**(避免 tombstone resurrection)
  - 本機 setRules([]) / clearAllRuleTombstones / clearAllAiMemory / clearRuleHistory
  - `pushNow('post-wipe', { wipeMarker: true })` 在 syncMeta 刻 marker
- normal push 保留 cloud 現有 wipeMarker(讓第三台機器加入時仍能 trip)
- 其他機器 `doPull` 偵測 `meta.wipeMarker.byMachineId !== ours && meta.wipeMarker.at > settings.lastSyncAt`:
  - 清本機 syncable 規則 + tombstones(在 `pullActive` critical section 內、防 echo push)
  - 持久化 `remoteWipeNotice` 到 `chrome.storage.local`
- Options 頁 `CrossMachineSyncCard` 加銀色 banner:「另一台機器執行了全部刪除、本機已自動同步清除 N 條、可從備份回滾」
- echo guard(`byMachineId !== ours`)+ staleness guard(`marker.at > lastSyncAt`)

### `dedupeRulesByKey` 啟動 migration

啟動時跑一次,把 (type, normalizeSignal, target) 三元組相同的規則合併(取最高 matchCount)。修復 import + sync 歷史殘留的重複規則。

### Stale 規則:從 disable 改為「直接刪除」

`autoDisableStaleRules` 中 stale 判定(`matchCount === 0` + createdAt > 100 天 OR `matchCount > 0` + lastUsedAt > 100 天)從「設 `enabled=false` + 寫 tombstone + 記 audit」改為**完全刪除**:
- 不留 tombstone(AI 可重新學)
- 不寫 audit
- 不寫 console
- 設計理由:長期沒命中等於不存在,留痕跡只汙染信任(用戶明確要求)

`high-error-rate`(`matchCount ≥ 20` + overrideCount/matchCount ≥ 0.5)仍走 soft disable + audit。

### autoUpgradeConflictRules:freshFolderPath

`autoUpgradeConflictRules`(衝突自動升級成 compound 的 handler)在 mint 新規則時、用 `folderById.get(effectiveFolderId).path` 取**當下**路徑,而非 `rule.targetFolderPath`。修復:資料夾被 rename 後升級出的 compound 規則仍帶 stale path。

OData filter PascalCase 正確化(`From/EmailAddress/Address`)+ server filter 失敗時 fall back 到 client-side filter,讓「點升級沒反應」的問題顯露 error。

### a11y polish (P2 audit)

- **Drawer focus trap + focus restoration**:`RuleDetailDrawer` 開啟自動 focus 關閉按鈕、Tab/Shift+Tab cycle within drawer、關閉 restore focus 到觸發 row
- **RuleRow 鍵盤啟用**:`tabIndex={0}` + Enter/Space → 開 drawer、`aria-label` 完整描述、`focus:ring-2 focus:ring-foreground/20 focus:ring-inset`
- **ActionsMenu menu role + arrow keys**:`role="menu"`、`role="menuitem"`、Arrow ↑↓ 導覽 / Home / End、Esc 關閉並 restore trigger focus
- **ConfirmDialog**:`role="alertdialog"`、Esc 取消、backdrop click 取消、`autoFocus` 在 Cancel(避免 Enter 誤確認危險動作)
- **RuleForm 狀態同步**:`useEffect([initial?.id])` reset state、切換規則時不殘留前一條的值;同一條規則的同步更新**不會 clobber 進行中編輯**(只有 id 變才 reset)

### Bug 掃描成果(4 agents 平行 audit)

並行掃過 UI / SW handlers / 規則引擎 / 同步、即時修復 7 項:
- Drawer Esc `stopImmediatePropagation` + `preventDefault`(防止 OWA FAB 攔截)
- ConfirmDialog Esc / backdrop click / autoFocus Cancel
- `chooseLearningSignal` P4/P5 衝突偵測用 `normalizeSignal`(原 `.toLowerCase()` 漏 `@`-prefix 情況)
- `requestEditRule` 設 `window.location.hash = 'rules-library/all'`(編輯按鈕觸發後正確導去全部規則)
- `wipeInFlight` guard(防 wipe 按鈕 double-click)
- `wipeAllRules` 順序:quiesce + clearCloud FIRST, then local setRules([])
- `bulkDelete` / `bulkDisable` / `bulkEnable` 用 `runBulk` helper + per-item try/catch、`bulkDelete` 預先關閉 drawer 若刪除的是當前打開的規則

### 新增測試

- `tests/sync-wipe-marker.test.ts` — wipeMarker 跨機器傳遞(8 tests):marker 刻寫 / 保留 / 替換 / 套用 / echo guard / staleness guard
- `tests/folder-activity-filter.test.ts` — 加 3 個 case:prefix 帶/不帶尾斜線、不誤匹 sibling

最終測試數:**395 tests across 22 test files**(舊版 367 tests / 21 files)。

---

## 2026-05-26 — 跨機器同步 v2(folderActivity)+ multi-machine visibility + onboarding wizard

### folderActivity 同步(schema v2)

- `SYNC_SCHEMA_VERSION` 從 1 升 2
- 同步「近日活動」最近 20 個 folder activity entries
- chunkFolderActivity helpers、`MAX_FOLDER_ACTIVITY_CHUNKS` = 4(defense-in-depth)
- v1 → v2 downgrade protection (Bug #H):cloud schemaVersion > ours 時 push refuse
- popup auto-refresh on open(data > 30 min)+ graceful display when folderId missing locally

### Multi-machine visibility(#6)

`SyncMeta.recentPushes[]`(cap 20、newest first):每次 push 附 `{ machineId, at }`,Options UI 顯示「A 推了 5 分鐘前 · B 推了 2 小時前」,讓使用者一眼掌握多機活動。

### Onboarding wizard(#3)

- `settings.onboardingDismissed`(per-device,Bug #T)
- `getOnboardingState` handler
- `OnboardingWizard` 元件:welcome → 分支(設 API key / 跑初始掃描 / 選 root folder)
- IdleScreen 條件性顯示(API key 缺 / 規則 0 / dismissed=false)

### Bug 修補一批

#L / #M / #N / #O / #P / #Q / #T / #U / #V — 詳細列表在 CLAUDE.md「已修 Sync Bug 一覽」。重要的:
- **#O**:concurrent `doPull` guard、`pullActive` check + set 必須同步
- **#Q**:`appendBackup` 加 mutex(RMW race)
- **#T**:`onboardingDismissed` 加進 `PER_DEVICE_SETTINGS_FIELDS`
- **#U**:`logError` 加 mutex
- **#V**:`recentPushes` `Array.isArray` 防 malformed payload

### 其他

- Claude API prompt caching(`extended-cache-ttl-2025-04-11`、ttl: '1h')
- ConversationMemory decay
- 中央 SW errorLog(200 entries cap)
- Sync error action suggestions(retry / upgrade / cleanup CTA)
- 拆檔:`src/background/handlers/sync.ts` 從 service-worker.ts extract

---

## 2026-05-23 — 跨機器同步 v1 + multi-bug audit

### 跨機器同步 v1

- 透過 `chrome.storage.sync` 同步:user_manual + ai_confirmed + ai_overridden 規則、tombstones(cap 500)、settings(扣 per-device)
- 5 個 backup snapshots rotation(`syncBackups`)
- onChanged listener install in `installSyncListener`(SW restart resilient)
- 5s debounce push + 200ms debounce remote-pull
- `PER_DEVICE_SETTINGS_FIELDS`:`claudeApiKey` / `syncMachineId` / `lastSyncAt` / `syncEnabled`

### 7 個 Critical / High Bug 修復(第一波)

- **#A**:per-device 欄位汙染(strip on push、filter on pull / restore)
- **#B**:first-enable 用 UNION 模式 merge(非 REPLACE,否則丟本機 unique 規則)
- **#C**:`restoreBackup` 保留 current `machineId` / `lastSyncAt` / `syncEnabled`
- **#D**:規則超過 chunk cap 時 UI surface truncation
- **#E**:remote-pull 改 UNION mode + tombstone-aware drop
- **#F**:`restoreBackup` reset tombstones(原本只 union、留下舊墓碑)
- **#G**:`pullInProgress` race fix(macrotask grace period)
- **#H**:`doPush` 拒絕 downgrade(cloud schemaVersion > ours)
- **#I**:`setSettings` mutex
- **#J**:Surface pull errors in status(原本只 console.warn)
- **#K**:`MAX_BACKUPS` 10 → 5(quota)

---

## 2026-05-22 — 通用化 + 全面 audit pass

### Breaking 通用化

- 硬編碼的事務所專屬資訊全部抽出到 `Settings`:`internalDomains` / `primaryRootPath` / `internalSubjectCategories`
- 舊使用者透過 silent migration 自動承接原值(example.com / 03 進行中案件 / 工時/薪資/利衝/行政/公告)
- Options 頁「事務所設定」更名為「**歸類偏好**」、內部信件規則改為 collapsed 子區塊、標明選用
- Classifier system prompt 中性化:「你是台灣律師事務所的郵件歸類助手」→「你是郵件歸類助手」;「案件」→「資料夾」;「案件代號」→「識別碼」
- `primaryRootPath` 輸入改用 `FolderPicker`、placeholder 不再洩漏特定資料夾名
- 全部 user-facing 文字統一用「內部」、移除「事務所 / 同事」字眼
- 新增 Settings JSON 匯出 / 匯入(不含 API key)

### 安全 / 隱私

- `aiIncludeFewShotExamples` 設定:可關閉「把規則 target path 當 few-shot 範例送 Claude」、保護資料夾名敏感的使用者
- Settings 匯入加版本檢查(`version: 1`)、未來版本格式不符會跳 confirm
- AI 自動偵測登入網域時、UI 明確顯示「偵測自 you@example.com」、避免多帳號搞錯

### Storage 退耗

- **`skipHistory`** 加 cap (5000) + 60 天 TTL,曾經會無限長
- **`folderActivity`** cap (200) 已存在、驗證
- **`subjectMemory`** conflictCount decay:連續 5 次同 folder 就 -1,意外 CC 不再永久毒化主旨

### 歸類管線改善

- **Thread Memory 不再 shadow `user_manual` / `case_code` / `compound`** 規則(以前手動建規則沒用,現在贏)
- **`bumpRuleHits` 加 errorRate gate**:`overrideCount/matchCount > 0.2` 時不升信心
- **`addRulesFilteringTombstones` mutex 內 dedup**:關掉並發 race
- **Sender 規則命中 From 時優於 domain 規則**(改 TYPE_PRIORITY 內順序)
- **`subject_keyword` bucket 按 signal.length 排序**:長 signal 先試、「112訴204」不再敗給「通知」
- **Unresolved → 律師手動填入 → 走 ai_overridden 學習流程**(以前完全沒學)
- **`generateAiConfirmedRules` / `Override` 共用 helper、 統一邏輯**

### 初始掃描

- **新增 generic-provider sender 抽取**(Pass A2):`andy@gmail.com` 等具體地址在某 folder ≥ 2 + 其他 folder ≤ 1 → 自動生 sender 規則
- **Pass A / A2 共用 `findUniqueFolderCandidate` helper**
- **掃描時 seed thread memory**:第一個 batch 就有 thread 覆蓋率、零額外 API

### Classifier 強化

- **AI 回傳路徑驗證**:`move` / `new_folder` 的目標 / 父資料夾不在樹中 → 自動降為 unresolved
- **Invalid JSON recovery**:解析失敗時嘗試 loose JSON(移除註解 / 尾逗號 / smart quotes)
- **Subject 截斷**:`MAX_SUBJECT_LEN = 200`、防超長 subject 拖累 prompt budget

### UX

- **Done screen `peekNextBatch`**:結束後輕量 API 探下一批、有 0 封就直接顯示「信箱已清空」、不再誤導
- **Popup status auto-poll**:OWA token 還在抓時自動每秒重抓 status(原本需關掉重開 popup)
- **Visibility / focus 自動 refresh**:popup 重新獲焦時更新狀態
- **近日活動 latestMessage 不再被 batch 沖掉**:`recordFolderActivityFromBatch` preserves prev.latestMessage;execute 也順手把當批 subject 寫入
- **近日活動拿掉今天/昨天/本週分段**:flat 列表、移除 +N badges
- **重新整理按鈕** 加 tooltip + 「10–30 秒」提示
- **Onboarding checklist**:首次使用顯示 3 步驟、完成第一個 batch 後自動消失
- **規則衝突 proactive 警告**:Rule editor 內、儲存前顯示衝突清單
- **`normalizeSubject` 擴充**:處理 Outlook 自動前綴(自動回覆 / Out of Office / [External] / 已讀回條 等)

### 文件

- 新增 README.md(使用者文件)
- 新增 CLAUDE.md(專案架構速查、給 AI 助手)
- 新增 CHANGELOG.md(本檔)

---

## 維護慣例

- 每次破壞 invariant 或加新 invariant 都要記到 `CLAUDE.md` 的「重要 invariants」段
- 每次新增 sync bug 修補要在 `tests/sync-engine.test.ts` 加對應 regression test、命名 `describe('Bug #X: ...')`
- 跨機器同步相關的變更都要過 `PER_DEVICE_SETTINGS_FIELDS` 審查
- 任何「靜默」行為(silent migration / silent cap / silent skip)都要在 UI 想辦法 surface,不然會變成「為什麼我的設定沒生效」的客訴源
