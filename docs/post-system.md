# 「貼網址 → 自動出 IG post」系統交接說明

呢份嘢寫俾第二個人（或者第二個 AI）睇，等佢可以照住起一套一模一樣嘅系統去對付第二啲網站。
實作喺 `post.html`、`gas/code.gs`、`vendor/qrcode.js`。

---

## 一、整體係四舊嘢串埋

```
瀏覽器頁面 (post.html，靜態，放喺 Vercel)
    ↓ ①攞商品資料           ↓ ②攞商品相
Google Apps Script (GAS) ←── 呢個係關鍵，係唯一去得到日本網站嘅一環
    ↓
瀏覽器用 canvas 畫出 1080×1350 商品卡
    ↓ ③上載成品圖
Cloudinary (unsigned upload preset)
    ↓ ④寫一行
Google Sheet「出Post」分頁
    ↓ ⑤讀一行出一個 post
Make.com → Instagram for Business
```

**冇伺服器、冇資料庫、冇 build step。** 全部係一個靜態 HTML ＋ 一個 GAS script。

---

## 二、點解一定要 GAS 做中間人（最重要嗰點）

瀏覽器**攞唔到**其他網站嘅 HTML —— CORS 擋死。所以：

- **商品資料**：頁面 `fetch(GAS + '?action=getXxxProduct&url=...')`，GAS 用
  `UrlFetchApp.fetch()` 攞返個 HTML，喺伺服器嗰邊 parse 完先回 JSON。
- **商品相**：瀏覽器直接 `drawImage()` 一張跨網站嘅相落 canvas，塊 canvas 會「被污染」，
  之後 `toBlob()` 會直接掟 exception，匯出唔到。所以要一個 `proxyImage` action ——
  GAS 攞返張相轉成 `data:image/jpeg;base64,...` 回俾前端，data URI 唔會污染 canvas。

GAS 仲有兩個附帶好處：

- **由 Google 個 IP 出去**，好多日本網站唔會擋
  （同一個網站，喺 GitHub Actions 度用 Playwright 爬係被擋死嘅）
- 內置 `LanguageApp.translate()` 免費做日→中翻譯

### GAS 嘅陷阱

`doGet` 係一連串 `if (action === ...)`，全部唔中會**靜靜跌落最尾嘅預設分流**，
回一份完全唔相干嘅嘢，呼叫方就會一直等到 timeout。
**一定要喺最尾加個擋位**，認唔到指令就直接報 `unknown_action`：

```javascript
if (action && action !== "getItems") {
  return jsonpOrJson_(param, {
    result: "error", error: "unknown_action", action: action,
    message: "呢個 GAS 版本未認得指令「" + action + "」，請重新部署。"
  });
}
```

仲有：GAS 改完要人手 **部署 → 管理部署 → ✏️ → 版本：新版本 → 部署**。
唔好撳「新增部署」—— 會換咗個 `/exec` 網址，所有呼叫方即刻死。

---

## 三、加一個新網站，只需要寫一個 function

GAS 入面每個來源一個 function，全部回**同一個形狀**：

```javascript
function fetchXxxProduct_(url) {
  if (!url || url.indexOf('xxx.com') === -1) return { error: '網址唔啱' };
  var resp = UrlFetchApp.fetch(url, {
    headers: {
      'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) ' +
                    'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
      'Accept-Language': 'ja,en;q=0.5'
    },
    muteHttpExceptions: true, followRedirects: true
  });
  if (resp.getResponseCode() !== 200) return { error: 'HTTP ' + resp.getResponseCode() };
  var html = resp.getContentText('UTF-8');

  var result = { name: '', image: null, price: null, variants: [], rawUrl: url };
  // …由 html 度抽資料（見下面「抽資料嘅優先次序」）…
  translateProductInPlace_(result);   // 日→中
  return result;
}
```

### 抽資料嘅優先次序（由最穩到最唔穩）

1. **`<script id="__NEXT_DATA__">`** — Next.js 網站成個商品物件都喺入面。
   唔好寫死 props 路徑（排行版同商品版路徑唔同），行勻棵 JSON tree 搵「似商品」嘅物件
   （有名 + 有圖 + 有價）。
2. **`<script type="application/ld+json">`** — JSON-LD `Product`，好多電商都有。
3. **`og:title` / `og:image`** — 幾乎一定有，最穩陣嘅 fallback。
4. **頁面入面嘅 JS 變數**，例如 `var itemName = "..."`
   ⚠️ 呢啲係 JS 字面值，入面 `（` 呢類轉義**未解碼**，要自己解，
   否則個名會出現成串 `（`：

   ```javascript
   function decodeJsString_(s) {
     return String(s)
       .replace(/\\u([0-9a-fA-F]{4})/g, function(_, h) {
         return String.fromCharCode(parseInt(h, 16));
       })
       .replace(/\\(["'\/\\])/g, '$1').trim();
   }
   ```
5. **DOM / regex 硬抽** — 最唔穩，網站改版就死。

### 抽價錢嘅陷阱（全部真係踩過）

- `<p>¥2,530<span>税込</span></p>` — 只睇葉節點會走漏，
  要睇每個元素嘅**直屬文字節點**（`n.nodeType === 3`）
- `¥1.980` — 有網站用**句點**做千位分隔，唔處理會變 ¥1
- 同一版有兩個價（批發價／建議零售價）—— 要搵個獨有嘅標記分開佢哋。
  例如日文批發站，零售價寫 `円/点（税抜）`、批發價寫 `円（税抜）`，靠 `円/点` 就分得出：

  ```javascript
  var rp = html.match(/([0-9][0-9,]*)\s*<span[^>]*class=["']taxUnit["'][^>]*>\s*円\s*\/\s*点/);
  ```

---

## 四、出卡係用 canvas 喺瀏覽器度畫

**唔好用 Canva**（手動，斷咗自動化），
**亦都唔好用 Cloudinary 文字疊加**（中文要另外上載字型，好易撞板）。

```javascript
const cv = document.createElement('canvas');
cv.width = 1080; cv.height = 1350;      // IG 4:5，feed 佔最高位置
const ctx = cv.getContext('2d');
// 畫底色、線、圖、字…
return await new Promise(ok => cv.toBlob(ok, 'image/jpeg', 0.92));
```

幾個一定要知嘅位：

- **字型要等佢載好先畫**，唔係會用咗 fallback 字：

  ```javascript
  await document.fonts.load('500 54px "Noto Serif TC"');
  await document.fonts.ready;
  ```
- **中文由瀏覽器自己 render，100% 準**，冇缺字問題（AI 出圖同 Cloudinary 都做唔到）
- **Instagram 個 API 只收 JPEG**，PNG 會被拒
- `ctx.letterSpacing = '6px'` 現代 Chrome 支援，可以做字距
- 圖片「填滿裁切」要自己計：
  `scale = Math.max(w / img.width, h / img.height)`，再置中
- 商品相要留返位俾兩行商品名 —— 相太高嘅話，第二行會壓住下面條線同頁腳

### QR code

用 `qrcode-generator`，直接攞 module 矩陣自己畫格仔落 canvas，唔經 DOM：

```javascript
const qr = qrcode(0, 'M');           // type 0 = 按長度自動揀版本
qr.addData(url); qr.make();
const n = qr.getModuleCount(), cell = size / n;
ctx.fillStyle = '#FFF';
ctx.fillRect(x - 8, y - 8, size + 16, size + 16);   // 四邊留白，唔留就掃唔到
ctx.fillStyle = '#000';
for (let r = 0; r < n; r++) for (let c = 0; c < n; c++) {
  if (qr.isDark(r, c)) ctx.fillRect(x + c * cell, y + r * cell, cell + 0.6, cell + 0.6);
}
```

`+0.6` 係避免縮放之後格與格之間見到白線。

**呢個 library 要放喺自己個 repo，唔好用 CDN** —— CDN 一失敗就會靜靜咁出一批冇 QR 嘅卡。
載唔到要直接掟錯，唔好靜靜地照出。

---

## 五、上載同存檔

### Cloudinary unsigned upload preset

Settings → Upload → Add upload preset → Signing mode: **Unsigned**

```javascript
const fd = new FormData();
fd.append('file', blob);
fd.append('upload_preset', 'your_preset_name');
const r = await fetch(`https://api.cloudinary.com/v1_1/${CLOUD}/image/upload`,
                      { method: 'POST', body: fd });
const url = (await r.json()).secure_url;
```

用 unsigned 就唔使將 API secret 放喺公開頁面。

### 寫入 Google Sheet（GAS `doPost`）

```javascript
fetch(GAS_URL, {
  method: 'POST',
  headers: { 'Content-Type': 'text/plain;charset=utf-8' },   // ← 關鍵
  body: JSON.stringify(payload)
});
```

用 `text/plain` 唔會觸發 CORS preflight，所以**讀得返 GAS 嘅回覆**。
用 `application/json` 就要 `mode:'no-cors'`，變成盲送，成功與否只能靠估。

### Sheet 一行 = 一個 post

```
A建立時間  B地區  C店舖  D件數  E商品名  F圖片網址(全部)
G商品網址  H文案  I Hashtag  J狀態  K出咗時間  L-U圖片1…圖片10
```

`圖片1/2/3` 逐張開一欄，唔係只擺一格用換行分隔 ——
Make 逐欄 map 好過喺佢度將一格拆做 array，少一步易錯嘅嘢。

---

## 六、Make 嗰邊（三個 module）

```
① Google Sheets · Search Rows
   Filter: 狀態(J) = 待出 ／ Order by 建立時間(A) asc ／ Limit 1

② Instagram for Business · Create a Carousel Post
   Photo 1/2/3 → 圖片1/2/3
   Caption → {{文案}} 空行 {{Hashtag}}

③ Google Sheets · Update a Cell ×2
   J{{Row number}} = 已出
   K{{Row number}} = {{formatDate(now; "YYYY-MM-DD HH:mm")}}
```

踩過嘅坑：

- `now` 要用**變數**（紫色藥丸），打字打「now」入去會寫個英文字入格
- Row number 要 map Search Rows 出嘅 `__ROW_NUMBER__`，唔可以填死數字
- Carousel **最少 2 張最多 10 張**。你 map 咗幾多格就要**每次都揀夠幾多件**，
  空 URL 會令成個 module 報錯
- 所以前端要**強制分組**：揀 8 件 ÷ 每個 post 3 件 = 存 2 行，
  剩低 2 件唔夠一組就唔存，並且話返用家知
- 用 `Update a Cell` 好過 `Update a Row`，唔會有誤清其他欄嘅風險
- 改 blueprint JSON 嗰陣，唔好改動 `restore.expect` 入面嘅 `mode`（`edit` vs `chose`）——
  改錯會 run 到一半彈 `Value not found in options`

---

## 七、其他實戰經驗（每個都係踩完先知）

**唔好爬列表頁。**
試過爬「人氣排行」「特價」呢啲貨架版，結果：一個網站個 bot 防護令 `page.goto`
等足 60 秒都唔返；另一個網站啲貨架係前端渲染，爬到得一件貨。
**但單一商品頁一直攞得到。**
所以改成由人揀貨貼網址，反而又快又穩。

**批次翻譯要有 fallback。**
將一件商品所有字串合併成一次 `LanguageApp.translate` 呼叫（快好多），
但 Google 有時會併行或拆行，令行數對唔返。如果對唔返就整批唔套用，
一件有幾十個尺碼顏色嘅衫就會**連商品名都唔譯**。
要改成：對唔返就逐個譯，慢啲但唔會全軍覆沒。

**日文商品名要斬短。**
日文商品名成句堆賣點，例如「＜大尺碼＞5/8 袖印花束腰T卹（吸濕排汗，快乾）（輕薄布料）」。
處理方法：

- 括號成組拎走（唔好喺第一個括號度斬，會淨返「5」）
- 只當 `｜` `・` 係分隔號（斜線唔算，「5/8 袖」係尺寸唔係兩個名）
- 斬剩少過 4 個字就當斬錯，用返原名

**最緊要係做成可以喺頁面手改** —— 自動化只係起點，唔係終點。

**唔好喺卡上面寫死匯率換算價。**
張卡會留喺 feed 度成年，匯率會變，寫死遲早唔啱。只出原幣價，本地價叫人 DM ——
順帶仲逼到人同你傾偈。
（例外：如果嗰個係你自己算好嘅賣價而唔係即時匯率換算，就出得。）

**品牌／角色名重複。**
商品名本身好多時已經帶住品牌或者角色名，前面再加一次就會變
「Kuromi Kuromi 三摺銀包」。前置之前要 check `name.includes(brand)`。

**地區唔止係換貨幣。**
香港叫 Kuromi、比卡超、多啦A夢；台灣叫酷洛米、皮卡丘、哆啦A夢。
字典要一隻公仔存兩個叫法。

**前端密碼閘要小心 CSS 特異度。**
`#lock { display:flex }` 會蓋過 `[hidden] { display:none }`，
塊遮罩變咗透明但仍然擋住所有掣。要寫 `#lock[hidden]{display:none}`。
另外要知：前端 PIN 擋得住路過嘅人，擋唔住肯睇 source code 嘅人。

**每次改動都要用真嘢測。**
全程用 Playwright 攔截 GAS／Cloudinary 嘅請求做假回覆，render 真張卡出嚟睇，
仲用 `jsQR` 解碼驗返個 QR 真係掃到正確網址。
九成 bug 係咁樣先揾到嘅 —— 淨係睇 code 睇唔出。

---

## 八、要換去第二個網站，改呢幾樣

| 改邊度 | 改乜 |
|---|---|
| GAS | 加一個 `fetchXxxProduct_()`，回上面嗰個形狀；喺 `doGet` 加個 `if (action === 'getXxxProduct')` |
| 前端 `SRC` 物件 | 加一行 `xxx: { label:'…', action:'getXxxProduct', re:/xxx\.com/i }` |
| 前端 `PAPER` 物件 | 加一套底色（底、線、編號、相框底要一齊轉色溫，唔係暖色線配冷色紙會好污糟） |
| `drawCard()` | 如果新來源要出唔同嘢（例如出角色名而唔係店名、出售價而唔係原幣價），加分支 |
| 文案 `buildCaption()` | 運費、出貨頻率、CTA 呢啲照抄改字 |

**其餘全部唔使動。**

---

## 九、有個來源唔可以出名（如果你都有呢個需要）

其中一個來源係批發站，個名一個字都唔可以出街。做法：

- 前端 `SRC` 個 `label` 直接寫成對外嘅代稱，成個系統只用嗰個 label
- 卡上面嗰行改為出**角色名**而唔係店名
- hashtag 用角色名，唔用店名
- 存落 Sheet 嘅「店舖」欄都係代稱

寫完之後 grep 成個檔案確認一次，剩返嘅只應該係內部識別碼
（物件 key、GAS action 名、網址判斷嘅 regex），嗰啲唔會出街。
