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
