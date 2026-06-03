# Mail Organizer

[![CI](https://github.com/ksliao0314/mail-organizer-extension/actions/workflows/ci.yml/badge.svg)](https://github.com/ksliao0314/mail-organizer-extension/actions/workflows/ci.yml)

Chrome MV3 擴充工具:讀取 Outlook 網頁版收件夾、用「既有規則 + Claude AI」把每封信歸類到對應資料夾。

設計鎖定的是「長時間、高郵件量」的工作流程——例如每天 50–100 封進的律師事務所、財務團隊、客服信箱。AI 不是唯一決策來源:規則命中為主、AI 補位、使用者每次的修正都會被學成新規則。

---

## 為什麼選這個工具(vs. 自己刻 Outlook Rules)

| 比較項 | Outlook 原生 Rules | Mail Organizer |
|---|---|---|
| 設定方式 | 一條一條手動寫 | AI 看完信自動分類,使用者改後系統自動學 |
| 主旨判斷 | 子字串(全字串才命中) | 全段主旨(正規化過,自動去 Re:/Fwd:) |
| 多條件 | 受限的 GUI | compound 規則(網域 AND 主旨關鍵字) |
| 衝突偵測 | 無 | 自動識別「同網域指向不同資料夾」 + 升級為 compound |
| 跨機器 | 透過 M365 同步,但卡住時無從查 | 主動同步,失敗會浮出 banner + 提供 rollback |
| 預覽 / 撤回 | 無預覽,沒有 undo | 跑前可逐封改 / 跑後 30 秒撤回 |
| 信心門檻 | 全部一視同仁 | 低信心自動跳「待決」、停在收件匣等人工處理 |

---

## 功能總覽

### 分類管線
- **規則優先 → AI 補位 → 待人決定** 三層
- 規則命中(case_code / compound / domain / subject_keyword / sender)直接歸檔
- 沒命中的送 Claude AI 評估
- AI 信心 < 門檻(預設 0.5)自動標「待決」、停在收件匣

### 學習機制(先廣後窄)
- AI 命中後 → 自動產生對應規則
- 使用者修正 AI → 走 `ai_overridden` 流程,下批由規則命中
- **先廣後窄**:同一個網域第一次看到 → 建 plain-domain 規則。再看到同網域去**另一個**資料夾 → 自動升級為 compound(網域 + 整段主旨)、舊規則自動沉默
- 對話延續:同一 Conversation thread / 同主旨後續信件,自動進同一個資料夾,**不打 AI**

### 規則生命週期
- **創建**:`chooseLearningSignal` 6 層 hierarchy(case_code / 法院字號 / 內部案號 / 網域 / 寄件人 / 內部網域 fallback)
- **休眠**:超過 100 天沒命中 → 直接刪除(不留痕跡)、`matchCount ≥ 20` 且錯誤率 ≥ 50% → 標記「自動休眠」、保留可重啟
- **復活**:同 type+signal+target 三元組再次出現 → 自動把休眠的規則重啟

### 規則庫 UI
- **儀表板**:總數 / 啟用 / 衝突 / 休眠 / 命中 Top 5,一眼掌握規則庫健康度
- **全螢幕工作區**:點「打開規則庫」進入,左側 6 個分頁(全部 / 衝突 / 休眠 / 健康 / 初始掃描 / 歷史)
- **按資料夾分組**(預設):同一個目標資料夾的規則收在一組,標題列顯示 N 條 / 命中總數 / 平均準確率
- **快速開關**:每條規則的 ⏻ 按鈕一鍵停用 / 啟用,不用進編輯頁
- **詳細抽屜**:點任一規則打開右側抽屜,看完整資料、編輯、刪除
- **批次操作**:勾選後一次停用 / 啟用 / 刪除

### 跨機器同步(可選)
- 透過瀏覽器原生帳號同步(Chrome → Google / Edge → Microsoft)同步規則、墓碑、設定
- 不同步:API key(永遠在本機)、auto_scan 規則(各機器各跑初始掃描)、本機快取
- **多機器衝突**:union 模式 merge、墓碑驅動的「上游刪除」傳遞
- **「全部刪除」會跨機器**:在 A 機按「全部刪除」、B 機下次同步會自動套用、銀色 banner 提示「另一台機器執行了全部刪除」、可從備份回滾
- **失敗顯紅 banner**:不會「靜默壞掉」、附帶具體 action(重試 / 升級 / 清理)

### 批次預覽 / 編輯 / 撤回
- 跑前可逐封改 action(歸檔 / 刪除 / 新建資料夾 / 略過)、改目標資料夾
- 跑完後 30 秒內可一鍵撤回(僅 move 操作、新建資料夾不還原)
- 略過的信件以 ID 記錄、60 天 / 5000 cap

### 初始掃描
- 第一次安裝時掃既有資料夾結構、為每個資料夾自動產出規則骨架
- **Pass A**(網域)→ **Pass A2**(generic-provider 完整地址寄件人)→ **Pass B**(主旨關鍵字)→ **Pass C**(案件代號)
- 同時 seed thread memory,第一批就有覆蓋率

### 規則匯出 / 匯入
- JSON 匯出(含墓碑、規則統計)、可在新裝置 / 跨團隊還原
- 匯入後信心會被當前系統 cap 調低時,UI 會明確提示「N 條規則的信心度降低、最大降幅 X.XX」、避免靜默變動

### 設定備份
- 匯出 / 匯入所有設定(不含 API key)、用於新裝置還原配置

---

## 安裝

```bash
npm install
npm run build      # 產出 dist/
```

Chrome / Edge → 擴充功能 → 開發人員模式 → 載入未封裝 → 選 `dist/`

---

## 第一次使用

擴充安裝後,popup 會顯示 onboarding wizard 帶你走過 3 步:

1. **設定 Claude API key**:從 [console.anthropic.com](https://console.anthropic.com/) 拿 key、貼進 Options 頁。永遠不會匯出、不會送 Outlook
2. **選擇主要根資料夾**(建議):Options 頁「歸類偏好」卡片用 FolderPicker 從你的實際資料夾結構選一個。被當成初始掃描的預設起點 + AI prompt 範例字首
3. **若你有工作信箱**(選用):展開「內部信件規則」、按「自動偵測」抓你登入的網域、避免同事的寄件人地址被學成獨立規則
4. **跑初始掃描**(建議):Options 頁「初始掃描」按開始、系統會掃你主要根資料夾下的所有子資料夾、自動生成規則骨架
5. **回 popup 按「開始歸類」** 跑第一批

---

## Options 頁配置

3 個分頁:**分類引擎** / **規則庫** / **資料 & 同步**。

### 分類引擎

| 欄位 | 預設 | 說明 |
|---|---|---|
| Claude API key | (空) | Anthropic key。**必填**。永遠本機儲存、永不匯出 |
| 分類模型 | claude-sonnet-4-6 | Claude model |
| 一批最多封 | 50 | 每次「開始歸類」抓最近 N 封信 |
| AI 信心門檻 | 0.5 | 低於此值的 AI 判斷標「待決」、不自動執行 |
| AI 是否參考既有規則範例 | 開 | 規則 target path 當 few-shot 送 Claude。資料夾名敏感時可關 |
| OWA 浮動圖示 | 開 | outlook.* 頁面右下顯示捷徑按鈕 |
| Pipeline 模式 | 關 | 完成一批時、背景預跑下一批 |

「歸類偏好」(同分頁底部):
- **主要根資料夾**:初始掃描起點 + AI prompt 範例字首
- **排除資料夾**:不會被分類進這些路徑下的子資料夾
- **內部網域**:同網域寄件人不被學成獨立規則、可按「自動偵測」用 `/me` API 抓
- **內部信件分類提示**:寫進 prompt 提示內部信怎麼分類

### 規則庫

預設顯示儀表板:
- 4 張卡片:總數 / 啟用 / 衝突 / 休眠
- Top 5 命中規則
- 「打開規則庫」按鈕進入全螢幕工作區

全螢幕工作區內 6 個子分頁:
- **全部**:預設按資料夾分組、可切換 flat / 搜尋 / 篩選類型 / 篩選來源
- **衝突**:同 signal 不同目標的規則對、可一鍵升級為 compound
- **休眠**:被自動休眠的規則、可手動啟用
- **健康**:錯誤率 / 孤兒 / 重複 等診斷
- **初始掃描**:重跑某個根資料夾的掃描
- **歷史**:規則變更 audit log(最近 500 筆)

### 資料 & 同步

- **跨機器同步卡片**:啟用 / 停用、Recent pushes(顯示哪台機器何時推了)、備份歷史(5 個 rotation)
- **匯出 / 匯入規則**(含墓碑)
- **匯出 / 匯入設定**(不含 API key)
- **重設 / 全部刪除**:不可復原。同步啟用時、其他機器下次同步會自動套用

「近日活動 顯示範圍」:用 FolderPicker 加路徑前綴 / 葉資料夾名,空白 = 顯示全部。

---

## 鍵盤捷徑(規則庫工作區)

| 鍵 | 動作 |
|---|---|
| `Tab` / `Shift+Tab` | 跳到下/上一個可聚焦元素 |
| `Enter` 或 `Space` | 在規則 row 上 → 開詳細抽屜 |
| `Esc` | 關抽屜 / 關確認對話框 / 關下拉選單 |
| `↑` `↓` | 在「操作」下拉選單內導覽 |
| `Home` `End` | 跳到選單第一/最後一項 |

---

## 開發

```bash
npm run dev          # vite dev mode + HMR
npm run type-check   # tsc --noEmit
npm test             # vitest run (395 tests)
npm run build        # 完整 build → dist/
npm run lint         # eslint
```

測試環境:vitest + jsdom + mocked chrome.storage(`tests/setup.ts`)。

---

## 主要 modules

| 路徑 | 內容 |
|---|---|
| `src/popup/` | toolbar popup(React + Tailwind + shadcn)、state machine + screens |
| `src/options/` | 設定 / 規則庫 / 規則健康度 / 同步面板 |
| `src/content/` | outlook.* 頁面注入(MSAL token 抓取 + FAB) |
| `src/background/` | service worker:訊息 router、initial-scan、execute、token cache、sync-engine、stale-sweep |
| `src/shared/` | 跨層共用(types / storage / classifier / rules / normalize / outlook-api / sync-chunks / browser-detect / folder-activity-filter) |
| `tests/` | vitest(395 tests across 22 files) |

---

## 隱私

- **API key 只存在 `chrome.storage.local`**、永遠不離開本機、不參與跨機器同步、不會匯出、不送 Outlook
- **AI 送到 Claude 的內容**:每封信的主旨(≤ 200 字截斷) + 寄件人 + 收件人 + 預覽(≤ 200 字) +(預設開,可關)你已驗證規則的 target path 當 few-shot 範例
- **MSAL token** 存在 `chrome.storage.session`(in-memory、瀏覽器重啟即清空)
- **跨機器同步只走帳號內**(Chrome via Google account / Edge via Microsoft account)、不通過任何第三方伺服器
- **不蒐集 / 不上傳任何使用統計**

---

## 限制

- 只支援 Outlook 網頁版(`outlook.office.com` / `office365.com` / `cloud.microsoft`),沒做桌面 Outlook
- 只支援單一 Outlook 帳號(以登入的 active tab 為準)
- 跨瀏覽器同步:Chrome 用 Google account、Edge 用 Microsoft account,**Chrome ↔ Edge 不自動同步**(可手動匯入匯出規則)
- chrome.storage.sync 上限 100 KB / 120 writes/min — 大規則庫(> 約 1500 條)或極端 burst 會觸發 truncation / quota,UI 會明確顯示
- Taiwan 法院字號 / 案件代號(`25A0067A`、`112訴204`)的 regex 跑在所有使用者身上、對非台灣案件自然 0 命中、無副作用

---

## License

Copyright © 2026 ksliao0314

本專案採 **GNU General Public License v3.0(或更新版本)** 授權 — 詳見 [`LICENSE`](LICENSE)。

你可以自由使用、研究、修改、散布本軟體;但**散布衍生作品時必須同樣以 GPL 開源並附上原始碼**(copyleft)。軟體按「現狀」提供,不負任何擔保責任。
