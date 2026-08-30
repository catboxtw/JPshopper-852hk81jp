# 「貼商品網址 → 自動出 IG post」系統：整體運作說明

呢份嘢講**成套系統點樣運作**，唔綁死喺任何一個購物網站。
睇完應該足夠由零起返一套，去對付你自己嗰啲網站。

一句講晒：**你貼幾條商品網址入去，撳一下，佢自動出好三張品牌商品卡、
寫好文案同 hashtag、上載好圖，然後排隊等自動出 post。**

---

## 一、整體係四舊嘢串埋

```
瀏覽器頁面（一個靜態 HTML，放喺 Vercel／Netlify／GitHub Pages 都得）
    ↓ ① 攞商品資料        ↓ ② 攞商品相
Google Apps Script（GAS）←── 關鍵：呢個係唯一去得到目標網站嘅一環
    ↓
瀏覽器用 <canvas> 畫出 1080×1350 商品卡
    ↓ ③ 上載成品圖
圖片寄存（Cloudinary，unsigned upload preset）
    ↓ ④ 寫一行
Google Sheet（一行 = 一個 post）
    ↓ ⑤ 讀一行，出一個 post
Make.com → Instagram for Business
```

**冇伺服器、冇資料庫、冇 build step。**
全部就係一個靜態 HTML ＋ 一個 Google Apps Script。成本近乎零。

---

## 二、五個步驟實際發生咩事

### ① 攞商品資料

用家喺頁面貼低商品網址（一行一條）。頁面逐條打去 GAS：

```
GET  https://script.google.com/macros/s/…/exec?action=getXxxProduct&url=<商品網址>
```

GAS 攞返個 HTML，parse 完回一個**固定形狀**嘅 JSON：

```json
{
  "name":  "商品名（已譯做中文）",
  "image": "商品相網址",
  "price": 1550,
  "variants": [ { "group": "顏色", "options": [ { "name": "M", "inStock": true } ] } ]
}
```

**每個網站寫一個 function，全部回同一個形狀。** 前端唔使知邊個網站點抽資料。

### ② 攞商品相

再打多次 GAS，佢攞返張相轉成 `data:image/jpeg;base64,…` 回俾前端。
（點解要咁做，見下面第三節。）

### ③ 畫卡

前端喺 `<canvas>` 度畫一張 1080×1350（IG 4:5）嘅商品卡：
底色、商品相、商品名、價錢、QR code、你嘅 handle。
畫完 `canvas.toBlob(…, 'image/jpeg')` 出圖。

### ④ 上載 ＋ 存檔

三張卡上載去圖片寄存攞返公開網址，再連同文案、hashtag 一齊寫一行入 Google Sheet，
狀態標「待出」。

### ⑤ 出 post

Make 每隔一陣讀一行「待出」，出 IG carousel，出完將狀態改「已出」。

---

## 三、點解一定要一個「中間人」（最重要嗰點）

**瀏覽器讀唔到其他網站嘅 HTML** —— CORS 擋死，冇得拗。
所以一定要有一個喺伺服器嗰邊行嘅嘢幫你攞。我用 Google Apps Script，因為：

- **免費、唔使養伺服器**、有一條固定嘅 `/exec` 網址
- **由 Google 個 IP 出去**，好多網站唔會擋
  （同一個網站，我用自己嘅爬蟲去爬係被 bot 防護擋死嘅，但 GAS 攞得到）
- 內置免費翻譯（`LanguageApp.translate`）
- 本身就住喺 Google Sheet 隔籬，寫資料入去零成本

**仲有一個好易忽略嘅原因：商品相都要經佢。**
如果你喺瀏覽器直接 `drawImage()` 一張跨網站嘅相落 canvas，
塊 canvas 會被標記為「已污染」，之後 `toBlob()` 會直接掟 exception，**匯出唔到**。
所以要由中間人攞返張相轉成 data URI —— data URI 唔會污染 canvas。

> 用第二樣嘢做中間人都得（Cloudflare Worker、Vercel Function、自己部 server），
> 原理一樣：**一個幫你攞 HTML 同圖片、回 JSON 嘅代理。**

### 中間人嘅陷阱

如果你好似 GAS 咁用一連串 `if (action === …)` 分流，
**全部唔中嗰陣一定要報錯**，唔好靜靜地跌落最尾嘅預設分支。
否則呼叫方會收到一份完全唔相干嘅資料，症狀係「好慢 ＋ 永遠載入中」，
查半日先發現原來只係未部署新版本。

---

## 四、抽資料嘅優先次序

同一個網站，用邊種方法抽差好遠。由最穩到最唔穩：

1. **頁面內嵌嘅 JSON**（`<script id="__NEXT_DATA__">`、`self.__next_f` 等等）
   現代網站成個商品物件都喺入面，仲齊過畫面顯示嗰啲。
   ⚠️ 唔好寫死 JSON 路徑（列表版同商品版路徑唔同），
   **行勻棵 tree 搵「似商品」嘅物件**（有名 + 有圖 + 有價）就最穩陣。
2. **JSON-LD**（`<script type="application/ld+json">` 入面嘅 `Product`）
3. **Open Graph meta**（`og:title` / `og:image`）—— 幾乎一定有，最好嘅 fallback
4. **頁面內嘅 JS 變數**（`var itemName = "…"`）
   ⚠️ 呢啲係 JS 字面值，`\uXXXX` 呢類轉義**未解碼**，要自己解，
   唔係個名會出現成串亂碼
5. **DOM／regex 硬抽** —— 最唔穩，網站改版就死

### 抽價錢特別多陷阱

- 價錢俾標籤斬開：`<p>¥2,530<span>税込</span></p>`
  只睇葉節點會走漏，要睇每個元素嘅**直屬文字節點**
- 千位分隔符可能係**句點**（`1.980` = 1980），唔處理會變 1
- 同一版有**幾個價**（原價／特價／批發價／建議零售價）
  要搵個獨有嘅標記分開佢哋，唔好靠位置
- 唔同尺碼唔同價 → 要出「起」字，唔係人哋照住最平嗰個價嚟問你就手尾長

---

## 五、點解用 canvas 畫卡

試過三條路，得一條行得通：

| 做法 | 結果 |
|---|---|
| AI 出圖 | 中文十次有七八次出錯，缺筆劃、錯別字 ❌ |
| 圖片服務疊字（Cloudinary 之類） | 中文要另外上載字型，好易撞板 ❌ |
| Canva 手動整 | 出到靚，但每個 post 要人手做，斷咗自動化 ❌ |
| **瀏覽器 `<canvas>`** | **中文由瀏覽器自己 render，100% 準，而且全自動** ✅ |

要注意：

- **字型要等載好先畫**，唔係會用咗 fallback 字：
  `await document.fonts.load('500 54px "字型名"'); await document.fonts.ready;`
- **Instagram 個 API 只收 JPEG**，PNG 會被拒
- 圖片「填滿裁切」要自己計 `scale = Math.max(w/img.width, h/img.height)` 再置中
- 文字要自己 wrap（逐個字試落去度闊度），同埋限定最多幾行
- 留位要諗定：商品名可能一行都可能兩行，唔留位第二行會壓住下面啲嘢

### QR code

用一個 QR library 攞返個 module 矩陣，自己畫格仔落 canvas，唔經 DOM。
每格畫大 0.6px，避免縮放之後格與格之間見到白線。四邊一定要留白，唔留就掃唔到。

**library 要放喺自己個 repo，唔好靠 CDN** ——
CDN 一失敗就會靜靜咁出一批冇 QR 嘅卡。載唔到要直接報錯。

---

## 六、上載同存檔

**圖片寄存**：Cloudinary 開一個 **unsigned upload preset**，
前端就唔使放 API secret 落公開頁面。

**寫 Google Sheet**：POST 去 GAS 嗰陣用 `Content-Type: text/plain;charset=utf-8`。
咁樣唔會觸發 CORS preflight，所以**讀得返 GAS 嘅回覆**，知道有冇成功。
用 `application/json` 就要 `mode:'no-cors'`，變成盲送，成功與否只能靠估。

**Sheet 設計成一行 = 一個 post**：

```
建立時間 │ 地區 │ 來源 │ 件數 │ 商品名 │ 商品網址 │ 文案 │ Hashtag │
狀態 │ 出咗時間 │ 圖片1 │ 圖片2 │ 圖片3 │ …
```

`圖片1/2/3` **逐張開一欄**，唔好只擺一格用換行分隔 ——
自動化工具逐欄 map 好過喺佢度將一格拆做 array，少一步易錯嘅嘢。

`狀態` 同 `出咗時間` 兩欄係俾自動化工具寫返嘅，用嚟避免同一個 post 出兩次。

---

## 七、自動出 post（Make 三個 module）

```
① Google Sheets · Search Rows
   Filter: 狀態 = 待出 ／ Order by 建立時間 asc ／ Limit 1

② Instagram for Business · Create a Carousel Post
   Photo 1/2/3 → 圖片1/2/3
   Caption → {{文案}} 空行 {{Hashtag}}

③ Google Sheets · Update a Cell ×2
   狀態 = 已出 ／ 出咗時間 = now
```

坑：

- **Carousel 最少 2 張最多 10 張**，而且你 map 咗幾多格就要**每次都夠幾多張**，
  空 URL 會令成個 module 報錯
- 所以**前端要強制分組**：揀 8 件 ÷ 每個 post 3 件 = 存 2 行，
  剩低嗰 2 件唔夠一組就唔存，並且話返用家知仲爭幾多件
- 更新狀態用 **Update a Cell** 好過 Update a Row，唔會誤清其他欄
- 時間要用**變數**，唔好打「now」呢個字入去（會照字面寫個英文字入格）
- 行號要 map 搜尋結果嗰個 row number，唔可以填死數字

---

## 八、幾個做完先知嘅道理

**唔好爬列表頁，叫人貼商品網址。**
我原本想自動爬「熱門商品」「特價」呢啲版做素材。結果：一個網站個 bot 防護
令頁面等足 60 秒都唔返；另一個網站啲貨架係前端渲染，爬到得一件貨。
**但單一商品頁一直攞得到，冇失敗過。**
改成由人揀貨貼網址，即刻又快又穩 —— 而且揀貨呢一步本身就應該有人把關。

**自動化只係起點，唔係終點。**
自動抽出嚟嘅商品名一定唔啱用（日文商品名成句堆賣點，又長又雜）。
所以每個欄位都要**做成可以喺頁面手改**，改完文案同張卡即刻跟住變。
呢個係整套嘢最實用嘅一點。

**唔好喺卡上面寫死匯率換算價。**
張卡會留喺 feed 度成年，匯率會變，寫死遲早唔啱，客人照住嚟問就尷尬。
只出原幣價，本地價叫人 DM —— 順帶仲逼到人同你傾偈。
（如果嗰個係你自己算好嘅賣價而唔係即時匯率換算，就出得。）

**留意名重複。**
商品名本身好多時已經帶住品牌名，你喺前面再加一次就會變
「XX牌 XX牌 三摺銀包」。前置之前要 check `name.includes(brand)`。

**地區唔止係換貨幣。**
同一件嘢，香港同台灣叫法可以完全唔同。要出兩個地區版就要諗埋用詞，
唔係淨係換個貨幣符號。

**前端密碼閘要小心 CSS 特異度。**
`#lock { display:flex }` 會蓋過 `[hidden] { display:none }`，
塊遮罩變咗透明但仍然擋住晒所有掣。
另外要知：**前端密碼擋得住路過嘅人，擋唔住肯睇 source code 嘅人。**

**每次改動都要用真嘢測。**
我全程用 Playwright 攔截外部請求做假回覆，render 真張卡出嚟睇，
仲用 QR 解碼器驗返個 QR 真係掃到正確網址。
**九成 bug 係咁樣先揾到嘅** —— 淨係睇 code 睇唔出。

---

## 九、要加一個新網站，改呢幾樣

| 改邊度 | 改乜 |
|---|---|
| 中間人（GAS） | 加一個抓取 function，回上面嗰個固定形狀；喺分流度加一個 action |
| 前端來源清單 | 加一行：對外顯示名、action 名、認網址嘅 regex |
| 前端配色表 | 加一套底色（底、線、編號、相框底要一齊轉色溫，暖色線配冷色紙會好污糟） |
| 畫卡 function | 如果新來源要出唔同嘢（例如出角色名而唔係店名），加分支 |
| 文案模板 | 運費、出貨頻率、CTA 呢啲改字 |

**其餘全部唔使動。**

---

## 十、如果有來源唔想出名

有時你唔想俾人知你去邊度入貨。做法：

- 前端嘅來源 `label` 直接寫成對外嘅代稱，**成個系統只用嗰個 label**
- 卡上面嗰行改為出第二樣嘢（例如角色名、系列名），唔出店名
- hashtag 用嗰樣嘢，唔用店名
- 連存落 Sheet 嗰欄都要用代稱

寫完 grep 成個檔案確認一次 —— 剩返嘅應該只有內部識別碼
（物件 key、action 名、認網址嘅 regex），嗰啲唔會出街。

---

## 十一、開頭要準備嘅嘢

1. **Google 帳戶** —— 開一個 Google Sheet，喺入面「擴充功能 → Apps Script」寫個 script，
   部署做「網頁應用程式」，任何人都可以存取。攞返條 `/exec` 網址。
   ⚠️ 之後每次改完 code 都要 **部署 → 管理部署 → ✏️ → 版本：新版本 → 部署**。
   **千祈唔好撳「新增部署」**，會換咗條網址，所有嘢即刻死。
2. **Cloudinary 免費戶口** —— 開一個 unsigned upload preset
3. **Make.com 免費戶口** —— 駁 Google Sheets 同 Instagram for Business
4. **IG 要係 Business／Creator 帳戶**，而且要連咗一個 Facebook 專頁，
   否則 API 出唔到 post
5. **一個放靜態 HTML 嘅地方** —— Vercel 免費戶口最方便
