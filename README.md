# KYX記帳

私人的錢和工作的錢（公司撥的製作費、交通費）分開算的記帳頁面。網頁放在 GitHub Pages，資料存在使用者自己的 Google 試算表（透過 Apps Script）。

- 記帳頁面：<https://harry87119.github.io/kyx/>
- 使用者設定教學：[給朋友的設定步驟.md](給朋友的設定步驟.md)
- 計算規則與交接說明：[HANDOFF.md](HANDOFF.md)

## 檔案

| 檔案 | 用途 |
|---|---|
| `index.html` | 記帳頁面（全部的畫面和計算都在這裡） |
| `apps-script/Code.gs` | 貼到使用者試算表裡的 Apps Script |
| `manifest.json`、`icons/` | 加入主畫面用的名稱和圖示 |
| `sw.js` | 離線快取，沒網路時也打得開頁面 |
| `.nojekyll` | 讓 GitHub Pages 直接提供檔案，不經過 Jekyll 處理 |

repo 裡不放任何個人記帳資料。

## 開啟 GitHub Pages

1. 把變更合併到 `main`。
2. 到 repo 的「Settings」→「Pages」。
3. 「Build and deployment」的 Source 選「Deploy from a branch」，Branch 選 `main`、資料夾選 `/ (root)`，按「Save」。
4. 等一兩分鐘，頁面會出現在 `https://harry87119.github.io/kyx/`。

免費方案的 GitHub Pages 需要 repo 是公開的。repo 裡只有程式碼，沒有任何人的記帳資料或通行碼，公開沒有問題。

## 更新頁面

改完 `index.html` 推到 `main` 就好。`sw.js` 會優先抓網路上的最新版，使用者下次打開就是新版本，不用另外處理快取。

## 資料格式

- 每筆紀錄：`{id, type, amount, note, date: "YYYY-MM-DD", ts, fx?}`（`fx`：浮動固定支出的實際金額對應哪一項）
- 設定：`{opening, until, pay, payday, pct, daily, notes, prefs, fixed: [{id, sid, name, amount, day, since: "YYYY-MM", end?: "YYYY-MM", vary?: true, from?: "YYYY-MM-DD", skip?: ["YYYY-MM-DD"]}]}`（pay＝手動改過的上次薪水，0＝自動抓；payday＝每月發薪日，預設 10；pct＝生活費比例，預設 35；daily＝每日基準，預設 500；notes＝通知狀態；prefs＝深淺色、明細檢視和排序、標籤順序、自己新增和刪掉的標籤）
- 瀏覽器裡的 `localStorage`：
  - `kyx-data-v1`：資料快取
  - `kyx-queue-v1`：待送清單
  - `kyx-sheet-v1`：試算表網址和通行碼
  - `kyx-serverv-v1`：試算表程式的版本（2＝可以只存一部分設定）
  - `kyx-notes-v1`、`kyx-theme`、`kyx-list-view`、`kyx-list-sort-v1`、`kyx-tag-order-v1`、`kyx-tag-custom-v1`、`kyx-tag-off-v1`：通知狀態、深淺色、明細檢視和排序、標籤順序、自己新增和刪掉的標籤（新版試算表程式也會同步）
- Apps Script 的 API：
  - `GET ?token=…`：回傳全部紀錄和設定。
  - `POST`（`Content-Type: text/plain`）：`{token, action}`，`action` 是 `add`（`entries` 陣列，id 重複會略過）、`delete`（`id`）或 `saveSettings`（`settings`，只寫有送來的欄位）。`GET` 回傳 `v: 2` 時，前端才會單獨送通知狀態和偏好。
