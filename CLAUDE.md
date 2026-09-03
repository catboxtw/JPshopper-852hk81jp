# JPshopper — 日本代購網站

## ⚠️ 永遠要做嘅嘢

**改咗 `gas/code.gs` 就一定要俾返 GitHub link，唔使等人問。**

```
https://github.com/catboxtw/JPshopper-852hk81jp/blob/main/gas/code.gs
```

Apps Script 冇得同 GitHub 同步，每次改完都要人手 copy 返落去，所以個 link 係必需品，唔係額外資料。

同時要提返部署步驟：

> Apps Script → **部署** → **管理部署** → ✏️ → 版本揀 **新版本** → 部署
> ⚠️ 唔好用「新增部署」—— 會換咗個 `/exec` 網址，成個站即刻死。

## 專案結構

| 檔案 | 做咩 |
|---|---|
| `gas/code.gs` | Google Apps Script 後端（`doGet` / `doPost` 按 `action` 分流）；改完要人手貼返上 Apps Script |
| `index.html` | 首頁 ＋ Nissen／ZOZOTOWN 長期代購即時報價 |
| `admin.html` | 後台（訂單、上架商品、購貨、匯率設定） |
| `post.html` | 出 post 助手（貼商品網址 → 商品卡 ＋ 文案）；喺 admin 入面，PIN 保護 |
| `shopping.html` | 客人選購頁（`/shopping-nissen`、`/shopping-zozo`），出過 post 嘅貨直接落單 |
| `shop-hk.html` / `shop-tw.html` | 落單前台（Alpine.js） |
| `flash.html` / `status.html` / `nissen-pay.html` | 搶購現貨、訂單狀態查詢、付款頁 |
| `tools/scrape-goods.mjs` | 日本商品清單抓取工具（Playwright） |
| `.github/workflows/scrape-goods.yml` | 喺 GitHub 上跑抓取，CSV 存入 `data/` |

## `doGet` 的陷阱

`doGet` 係一連串 `if (action === ...)`，全部唔中會跌落最尾嘅商品清單分流。加新 action 之後**唔重新部署**，呼叫方就會收到一份完全唔相干嘅商品清單，症狀係「好慢 ＋ 永遠載入中」。`gas/code.gs` 已經加咗擋位會直接報 `unknown_action`，但擋位本身都要部署先生效。

## 資料

- Google Sheet：訂單紀錄、購貨紀錄、盈利紀錄、易寄取地址、收單截止時間，加每個活動一個分頁
- Supabase：`app_settings`（匯率）、活動同訂單同步
- Cloudinary：商品圖片上傳

## 匯率：兩個站共用，Mercari 嗰邊定，JPshopper 跟

匯率**每兩星期改一次**，源頭喺另一個 repo `chowmichelle910-byte/852hk-81jp-mercari`
（另一個 session，同呢個 session 唔連通 —— 跨 owner 加唔到 repo）。
每次嗰邊攞到新匯率，**呢邊都要一齊更新**，Nissen 就係靠佢計價。

存喺 Supabase `app_settings` 表，一行一個 key：

| key | 做咩 |
|---|---|
| `rate_hk` | ¥1 = HK$？（稅込価格換算） |
| `rate_tw` | ¥1 = NT$？ |
| `ship_hk_50g` / `ship_tw_50g` | 運費，唔關匯率事 |

**改完唔使部署、唔使 build。** `index.html`、`post.html`、`shopping.html`
都係一載入就即場讀呢個表，所以改完人客重新整理就已經係新價。

兩個改法：

1. **後台**：`admin.html` → ⚙️ 匯率設定 → 填 → 儲存（正常用呢個）
2. **直接寫**（例如由 Mercari 嗰邊自動同步）：

```
POST {SUPA_URL}/rest/v1/app_settings?on_conflict=key
Header: apikey / Authorization: Bearer {anon key}
        Content-Type: application/json
        Prefer: resolution=merge-duplicates,return=minimal
Body:   [{"key":"rate_hk","value":"0.057","updated_at":"<ISO>"}]
```

`SUPA_URL` 同 anon key 喺 `admin.html` 頂頭（`SUPA_URL` / `SUPA_KEY`）。
`value` 係 **TEXT**，唔係數字，要 `String(n)`。

⚠️ `app_settings` 條 RLS policy 而家係 `USING (true) WITH CHECK (true)`，
即係**任何人揭開個網頁原始碼攞到 anon key 就改到匯率**。多一個系統寫入之前
知咗佢會好啲 —— 唔好將呢個表當成有保護。
