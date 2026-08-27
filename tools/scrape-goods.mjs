#!/usr/bin/env node
/**
 * 日本商品頁抓取工具 — 輸出可直接貼入 Google Sheet 的 CSV
 * ============================================================
 *
 * 用法（喺電腦 terminal 行）：
 *
 *   cd JPshopper-852hk81jp
 *   npm install playwright          # 第一次先要裝
 *   npx playwright install chromium # 第一次先要裝
 *
 *   node tools/scrape-goods.mjs "https://chiikawapark-tokyo.jp/goods/"
 *
 * 常用選項：
 *   --venue-only          ⭐ 剔走「🌐網購限定」商品，只留親身去到會場買得到嘅
 *   --strict-venue        再嚴格啲：淨係要明確標住「🏬會場限定」嗰啲
 *   --expect 400          ⭐ 預期最少幾多件；唔夠會出警告（防止頁面冇載入齊）
 *   --out goods.csv       輸出檔名（預設 goods.csv）
 *   --rate 0.22           日幣→台幣匯率（預設 0.22）
 *   --debug               抓唔到嘢時用：印出頁面結構樣本，send 俾 Claude 睇
 *   --static              唔用瀏覽器，直接抓 HTML（快，但 JS 載入嘅頁面會抓唔到）
 *   --max-scroll 80       最多向下捲幾多次（預設 80；捲到冇新嘢會自動停）
 *
 * 會出兩個檔案：
 *   goods.csv         欄位＝Google Sheet 商品分頁 A~J，由 A2 貼落去即可。
 *                     重量（G欄）留空要你自己填，填完 D/E/I 嘅公式會自動計價。
 *   goods-review.csv  畀你自己睇：標籤獨立一欄，方便 filter／排序。
 *
 * ⚠️ 限定標籤係靠掃描頁面文字認出嚟（オンライン限定 / パーク限定 等）。
 *    如果個網站用圖片 icon 而唔係文字標示，就捉唔到，行 --debug 傳俾 Claude 加。
 */

import { writeFileSync } from 'node:fs';

// ── 解析指令參數 ──────────────────────────────────────────
const argv = process.argv.slice(2);
const urls = argv.filter(a => /^https?:\/\//.test(a));
const flag = (name, def = null) => {
  const i = argv.indexOf('--' + name);
  return i >= 0 ? (argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[i + 1] : true) : def;
};
const OUT          = flag('out', 'goods.csv');
const TWD_RATE     = parseFloat(flag('rate', '0.22'));
const DEBUG        = !!flag('debug', false);
const STATIC       = !!flag('static', false);
const MAX_SCROLL   = parseInt(flag('max-scroll', '80'));
const EXPECT       = parseInt(flag('expect', '0')) || 0;
const STRICT_VENUE = !!flag('strict-venue', false);
const VENUE_ONLY   = !!flag('venue-only', false) || STRICT_VENUE;
const CHROME_PATH  = flag('chrome-path', process.env.CHROME_PATH || '');

if (!urls.length) {
  console.error('❌ 請提供至少一條網址，例如：\n   node tools/scrape-goods.mjs "https://chiikawapark-tokyo.jp/goods/"');
  process.exit(1);
}

// ── 喺瀏覽器入面行嘅抓取邏輯 ────────────────────────────────
// 原理：先搵出所有「顯示緊價錢」嘅元素，再向上搵返包住佢、
// 而且有兄弟姊妹（即係重複出現嘅商品卡）嘅容器，然後由卡入面
// 抽商品名 / 圖片 / 連結。咁樣唔使預先知道個網站嘅 class 名。
function extractInPage(debugMode) {
  const PRICE_RE = /(?:[¥￥]\s*([0-9][0-9,]*)|([0-9][0-9,]*)\s*円)/;

  // 商品標籤 — ⚠️ 「限定」同「対象」意思差好遠，一定要分清楚：
  //   オンライン限定      → 淨係網上有，現場買唔到       （--venue-only 會剔走）
  //   オンライン販売対象  → 呢件貨都有得網上賣，現場一般都有（保留）
  //   オンライン販売対象外 → 唔上網賣，即係淨係現場有     （保留，等同會場限定）
  const TAG_PATTERNS = [
    // 先判斷「対象外」，唔好俾下面條「対象」搶咗去
    { re: /オンライン(ストア|ショップ)?(販売)?対象外|店頭(販売)?のみ|会場(販売)?のみ|オンライン(販売)?(不可|なし)/,
                                                                                          tag: '🏬會場限定' },
    { re: /オンライン(ストア|ショップ)?限定|ONLINE\s?限定|WEB\s?限定|ウェブ限定|通販限定/i, tag: '🌐網購限定' },
    // 「販売対象」＝都有得網上賣，唔代表現場冇，所以淨係做參考標示
    { re: /オンライン(ストア|ショップ)?(販売)?対象(?!外)|オンライン(ストア|ショップ)?(でも)?(取扱|販売中)/,
                                                                                          tag: '🛒網上都有' },
    { re: /パーク限定|会場限定|店舗限定|館内限定|ここでしか|現地限定/,                      tag: '🏬會場限定' },
    { re: /予約(商品|受付|販売)|受注生産|お取り寄せ/,                                      tag: '預訂商品' },
    { re: /SOLD\s?OUT|完売|品切れ|在庫(なし|切れ)/i,                                       tag: '⛔售罄' },
    { re: /数量限定/,                                                                      tag: '數量限定' },
    { re: /期間限定/,                                                                      tag: '期間限定' },
    { re: /新商品|NEW\s?(ITEM|ARRIVAL)|新発売/i,                                           tag: '新商品' },
  ];
  // 每人限購：「お一人様◯点まで」→ 直接填入試算表 J 欄
  const LIMIT_RE = /お一人様[^0-9０-９]{0,6}([0-9０-９]+)\s*(?:点|個|コ|つ)/;
  const toHalfWidth = s => (s || '').replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));

  const parsePrice = (txt) => {
    const m = (txt || '').match(PRICE_RE);
    if (!m) return null;
    const n = parseInt((m[1] || m[2] || '').replace(/,/g, ''), 10);
    return Number.isFinite(n) && n > 0 ? n : null;
  };

  // 1) 搵所有「自己直接寫住價錢」嘅元素
  //    用直接文字節點嚟判斷，咁 <p>¥2,530<span>税込</span></p> 呢種都捉到
  const hasOwnPriceText = (el) => {
    for (const n of el.childNodes) {
      if (n.nodeType === 3 && PRICE_RE.test(n.textContent)) return true;
    }
    return false;
  };
  const leaves = Array.from(document.querySelectorAll('body *')).filter(hasOwnPriceText);

  // 2) 每個價錢向上爬，搵最細嘅「商品卡」
  const cards = new Set();
  for (const leaf of leaves) {
    let node = leaf;
    for (let depth = 0; depth < 8 && node.parentElement; depth++) {
      node = node.parentElement;
      const hasImg = !!node.querySelector('img');
      const sibs = node.parentElement
        ? Array.from(node.parentElement.children).filter(s => s.tagName === node.tagName)
        : [];
      // 有圖 + 有同類兄弟（代表係重複嘅列表項）= 應該就係商品卡
      if (hasImg && sibs.length >= 2) { cards.add(node); break; }
    }
  }

  // 3) 由每張卡抽資料
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const cardArr = Array.from(cards);

  // 預先收集「區塊標題」：排除喺商品卡入面嗰啲（因為卡入面個名可能都係 h4）
  const sectionHeadings = Array.from(
    document.querySelectorAll('h1,h2,h3,h4,[class*="heading" i],[class*="category" i]')
  ).filter(h => !cardArr.some(c => c.contains(h)));

  const out = [];
  for (const card of cardArr) {
    const price = parsePrice(card.textContent);
    if (!price) continue;

    // 商品名：優先揀 class/tag 睇落似標題嘅，否則揀最長嗰段非價錢文字
    let name = '';
    const titleEl = card.querySelector(
      '[class*="name" i],[class*="title" i],[class*="ttl" i],h2,h3,h4,figcaption'
    );
    if (titleEl && clean(titleEl.textContent) && !PRICE_RE.test(clean(titleEl.textContent))) {
      name = clean(titleEl.textContent);
    }
    if (!name) {
      const texts = Array.from(card.querySelectorAll('*'))
        .filter(el => el.children.length === 0)
        .map(el => clean(el.textContent))
        .filter(t => t && t.length >= 2 && !PRICE_RE.test(t) && !/^(税込|税抜|NEW|SOLD ?OUT)$/i.test(t));
      name = texts.sort((a, b) => b.length - a.length)[0] || '';
    }
    if (!name) continue;

    const img = card.querySelector('img');
    const imgUrl = img
      ? (img.currentSrc || img.src ||
         img.getAttribute('data-src') || img.getAttribute('data-original') || '')
      : '';

    const a = card.matches('a') ? card : card.querySelector('a');
    const link = a ? a.href : '';

    // 分類：喺文件順序上，排喺呢張卡之前嘅最後一個區塊標題
    let category = '';
    for (const h of sectionHeadings) {
      // 4 = DOCUMENT_POSITION_FOLLOWING（即 card 排喺 h 之後）
      if (h.compareDocumentPosition(card) & 4) category = clean(h.textContent).slice(0, 40);
      else break;
    }

    // 標籤 + 每人限購：卡入面同所屬分類標題都掃一次
    const scanText = clean(card.textContent) + ' ' + category;
    // 去重：🏬會場限定 有兩條 pattern（対象外 / パーク限定）會夾到同一件貨
    const tags = [...new Set(TAG_PATTERNS.filter(t => t.re.test(scanText)).map(t => t.tag))];
    const lm = scanText.match(LIMIT_RE);
    const limitPerPerson = lm ? parseInt(toHalfWidth(lm[1]), 10) || '' : '';

    out.push({ name, price, imgUrl, link, category, tags, limitPerPerson });
  }

  // 4) 去重（同名同價當作同一件）
  const seen = new Set();
  const products = out.filter(p => {
    const k = p.name + '|' + p.price;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });

  const debug = debugMode ? {
    priceLeafCount: leaves.length,
    cardCount: cards.size,
    sampleCards: Array.from(cards).slice(0, 3).map(c => c.outerHTML.slice(0, 900)),
  } : null;

  return { products, debug };
}

// ── CSV 產生（欄位順序 = Google Sheet A~J）─────────────────
const csvCell = v => {
  const s = String(v ?? '');
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
};

function toCsv(products, rate) {
  const header = [
    '商品名稱', '款式', '日幣原價', '成本HKD', '售價HKD',
    '圖片', '重量g', '備註', '台幣售價', '限購',
  ];
  const lines = [header.map(csvCell).join(',')];

  products.forEach((p, i) => {
    const r = i + 2; // 試算表實際行號（第 1 行係標題）
    lines.push([
      p.name,
      '',                                   // B 款式：留空，你自己填
      p.price,                              // C 日幣原價
      `=C${r}*0.05+G${r}/50*4.5`,           // D 成本HKD（自動計）
      `=CEILING(D${r}+30,1)`,               // E 售價HKD = 成本+30
      p.imgUrl,                             // F 圖片
      '',                                   // G 重量g：⚠️ 要你自己填
      [p.category ? '分類：' + p.category : '', ...(p.tags || [])]
        .filter(Boolean).join(' ／ '),      // H 備註（含分類＋限定標籤）
      `=CEILING(C${r}*${rate},5)`,          // I 台幣售價（進位至5的倍數）
      p.limitPerPerson || '',               // J 限購（お一人様◯点まで）
    ].map(csvCell).join(','));
  });

  return '﻿' + lines.join('\r\n') + '\r\n'; // BOM：確保 Excel 開日文唔會亂碼；結尾換行方便 wc -l 點算
}

// 第二個檔案：畀你自己篩選／排序用（標籤獨立一欄，方便 filter 網購限定）
function toReviewCsv(products) {
  const header = ['商品名稱', '日幣原價', '標籤', '分類', '每人限購', '商品連結', '圖片'];
  const lines = [header.map(csvCell).join(',')];
  products.forEach(p => {
    lines.push([
      p.name, p.price, (p.tags || []).join(' ／ '),
      p.category || '', p.limitPerPerson || '', p.link || '', p.imgUrl || '',
    ].map(csvCell).join(','));
  });
  return '﻿' + lines.join('\r\n') + '\r\n';
}

// ── 主流程 ────────────────────────────────────────────────
async function scrapeStatic(url) {
  console.log(`  → 直接抓 HTML（--static）…`);
  const res = await fetch(url, {
    headers: { 'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/120 Safari/537.36' },
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const html = await res.text();

  const { JSDOM } = await import('jsdom').catch(() => {
    throw new Error('--static 模式需要 jsdom，請先行：npm install jsdom');
  });
  const dom = new JSDOM(html, { url });
  global.document = dom.window.document;
  return extractInPage(DEBUG);
}

async function scrapeBrowser(url) {
  const { chromium } = await import('playwright').catch(() => {
    throw new Error('搵唔到 playwright，請先行：npm install playwright && npx playwright install chromium');
  });

  // --chrome-path：如果 npx playwright install 裝唔到，可以指定現成嘅 Chrome/Chromium
  const browser = await chromium.launch(
    CHROME_PATH ? { executablePath: CHROME_PATH } : {}
  );
  // 扮返個正常日本訪客：預設 headless UA 會寫住 HeadlessChrome，有啲網站會擋
  const page = await browser.newPage({
    viewport:   { width: 1440, height: 900 },
    userAgent:  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    locale:     'ja-JP',
    timezoneId: 'Asia/Tokyo',
    extraHTTPHeaders: { 'Accept-Language': 'ja-JP,ja;q=0.9' },
  });
  try {
    console.log('  → 開緊頁面…');
    await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });

    // 捲到底 + 撳「もっと見る」之類嘅載入更多按鈕
    // 一路捲一路報住而家有幾多件，等你睇到係咪仲載緊
    console.log('  → 捲動載入全部商品…');
    const countImgs = () => page.evaluate(() => document.querySelectorAll('img').length);
    let stable = 0, i = 0;
    for (; i < MAX_SCROLL && stable < 3; i++) {
      const before = await page.evaluate(() => document.body.scrollHeight);

      const moreBtn = page.locator(
        'text=/もっと見る|もっとみる|さらに表示|次へ|Load more|See more/i'
      ).first();
      if (await moreBtn.count() && await moreBtn.isVisible().catch(() => false)) {
        await moreBtn.click().catch(() => {});
        await page.waitForTimeout(1500);
      }

      await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
      await page.waitForTimeout(1200);

      const after = await page.evaluate(() => document.body.scrollHeight);
      stable = after === before ? stable + 1 : 0;
      if (i % 5 === 4) console.log(`     …捲咗 ${i + 1} 次，頁面已有約 ${await countImgs()} 張圖`);
    }
    // 用盡捲動次數但頁面仲喺度長 = 好可能未載齊
    if (i >= MAX_SCROLL && stable < 3) {
      console.warn(`  ⚠️  捲到上限 ${MAX_SCROLL} 次頁面仍然喺度加長，可能未載齊！`);
      console.warn(`      建議加大：--max-scroll ${MAX_SCROLL * 2}`);
    }

    // 觸發 lazy-load 圖片：由頭慢慢捲返落去
    await page.evaluate(async () => {
      for (let y = 0; y < document.body.scrollHeight; y += 600) {
        window.scrollTo(0, y);
        await new Promise(r => setTimeout(r, 60));
      }
    });
    await page.waitForTimeout(800);

    console.log('  → 抽取商品資料…');
    return await page.evaluate(extractInPage, DEBUG);
  } finally {
    await browser.close();
  }
}

(async () => {
  const all = [];
  const debugDumps = [];

  for (const url of urls) {
    console.log(`\n📦 ${url}`);
    try {
      const { products, debug } = STATIC ? await scrapeStatic(url) : await scrapeBrowser(url);
      console.log(`  ✅ 抓到 ${products.length} 件商品`);
      all.push(...products);
      if (debug) debugDumps.push({ url, ...debug });
    } catch (e) {
      console.error(`  ❌ 失敗：${e.message}`);
    }
  }

  // 跨頁去重
  const seen = new Set();
  let products = all.filter(p => {
    const k = p.name + '|' + p.price;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });

  // ── 只留會場買得到嘅商品 ──────────────────────────────────
  // ⚠️ 只有「オンライン限定」先算現場買唔到；「オンライン販売対象」
  //    只係話佢都有得網上賣，現場一般照樣有，所以唔會剔走。
  const isOnlineOnly = p => (p.tags || []).includes('🌐網購限定');
  const isVenue      = p => (p.tags || []).includes('🏬會場限定');
  const alsoOnline   = p => (p.tags || []).includes('🛒網上都有');

  if (VENUE_ONLY && products.length) {
    const before   = products.length;
    const online   = products.filter(isOnlineOnly).length;
    const venue    = products.filter(p => isVenue(p) && !isOnlineOnly(p)).length;
    const both     = products.filter(p => alsoOnline(p) && !isVenue(p) && !isOnlineOnly(p)).length;
    const untagged = products.filter(p => !isVenue(p) && !isOnlineOnly(p) && !alsoOnline(p)).length;

    products = STRICT_VENUE
      ? products.filter(p => isVenue(p) && !isOnlineOnly(p))   // 淨係要明確標住會場限定
      : products.filter(p => !isOnlineOnly(p));                // 只剔走「網購限定」，其餘照留

    console.log(`\n🏬 只保留會場買得到嘅商品${STRICT_VENUE ? '（嚴格模式）' : ''}：`);
    console.log(`     🏬 會場限定／唔上網賣　${venue} 件　→ 保留`);
    console.log(`     🛒 網上都有（販売対象）${both} 件　→ ${STRICT_VENUE ? '剔走（嚴格模式）' : '保留（現場一般都有）'}`);
    console.log(`     ⬜ 冇任何標示　　　　　${untagged} 件　→ ${STRICT_VENUE ? '剔走（嚴格模式）' : '保留（喺場一般買到）'}`);
    console.log(`     🌐 網購限定（唯一剔走）${online} 件　→ 剔走`);
    console.log(`     ${before} 件 → 保留 ${products.length} 件`);

    if (STRICT_VENUE && (both || untagged)) {
      console.log(`     ⚠️  嚴格模式剔走咗 ${both + untagged} 件冇明確標「会場限定」嘅貨。`);
      console.log(`         如果個網站係用「オンライン販売対象」嚟標示（即係反過嚟講），`);
      console.log(`         咁冇標嗰啲先至係現場限定，唔應該剔 —— 除去 --strict-venue 再行。`);
    } else if (!STRICT_VENUE && venue) {
      console.log(`     💡 想淨係要明確標住「会場限定」嗰 ${venue} 件，加 --strict-venue。`);
    }
  }

  if (!products.length) {
    if (all.length) {
      // 有抓到嘢，只係全部俾篩選規則剔走咗
      console.error('\n❌ 篩選之後一件都唔剩。');
      console.error(STRICT_VENUE
        ? '   --strict-venue 淨係要明確標住「会場限定 / パーク限定」嘅商品，但一件都認唔到。\n   除去 --strict-venue 再行一次（改為只剔走網購限定）。'
        : '   全部商品都被當成網購限定 —— 好可能係認錯，行 --debug 傳返俾 Claude 睇。');
      process.exit(1);
    }
    console.error('\n❌ 一件商品都抓唔到。');
    console.error('   請加 --debug 再行一次，然後將印出嚟嘅內容 send 俾 Claude 幫你調整。');
    if (debugDumps.length) {
      writeFileSync('debug-dump.json', JSON.stringify(debugDumps, null, 2), 'utf8');
      console.error('   （已寫入 debug-dump.json）');
    }
    process.exit(1);
  }

  const REVIEW_OUT = OUT.replace(/\.csv$/i, '') + '-review.csv';
  writeFileSync(OUT, toCsv(products, TWD_RATE), 'utf8');
  writeFileSync(REVIEW_OUT, toReviewCsv(products), 'utf8');

  console.log(`\n🎉 完成！共 ${products.length} 件商品`);
  console.log(`   價錢範圍：¥${Math.min(...products.map(p => p.price)).toLocaleString()} ~ ¥${Math.max(...products.map(p => p.price)).toLocaleString()}`);

  // 件數對唔對得上你嘅預期？唔夠嘅話大聲提你，唔好靜靜哋少咗貨
  if (EXPECT) {
    const scraped = all.length; // 篩選之前抓到嘅總數
    if (scraped < EXPECT) {
      console.warn(`\n   ⚠️⚠️  只抓到 ${scraped} 件，少過你預期嘅 ${EXPECT} 件！`);
      console.warn(`         個頁面好可能未載齊，唔好就咁攞去報價。試吓：`);
      console.warn(`         1. 加大捲動次數：--max-scroll ${MAX_SCROLL * 2}`);
      console.warn(`         2. 睇吓啲貨係咪分咗幾個分類頁，逐條網址一次過傳入：`);
      console.warn(`            node tools/scrape-goods.mjs "網址1" "網址2" "網址3" …`);
      console.warn(`         3. 仲係唔掂就行 --debug，傳 debug-dump.json 俾 Claude`);
    } else {
      console.log(`   ✅ 抓到 ${scraped} 件，已達到你預期嘅 ${EXPECT} 件`);
    }
  }

  // 標籤統計：一眼睇清有幾多件係網購限定 / 會場限定
  const tally = {};
  products.forEach(p => (p.tags || []).forEach(t => { tally[t] = (tally[t] || 0) + 1; }));
  const tagged = Object.entries(tally).sort((a, b) => b[1] - a[1]);
  if (tagged.length) {
    console.log(`\n   🏷️  標籤統計：`);
    tagged.forEach(([t, n]) => console.log(`        ${t}　${n} 件`));
  } else {
    console.log(`\n   ⚠️  一個限定標籤都捉唔到 —— 可能個網站唔係用文字標示（例如用圖片 icon），`);
    console.log(`        或者用咗其他字眼。行 --debug 傳返 debug-dump.json 俾 Claude 幫你加。`);
  }
  const noLimitInfo = products.filter(p => !(p.tags || []).some(t => /限定/.test(t))).length;
  if (tagged.length && noLimitInfo) {
    console.log(`        （另有 ${noLimitInfo} 件冇任何限定標示，落單前建議自己 double check）`);
  }

  console.log(`\n   📄 ${OUT}`);
  console.log(`      欄位＝Google Sheet A~J，由 A2 貼落去；填好「重量g」欄後成本／售價自動計。`);
  console.log(`   📄 ${REVIEW_OUT}`);
  console.log(`      標籤獨立一欄，用嚟 filter「🌐網購限定」／「🏬會場限定」睇邊啲買到。`);

  if (debugDumps.length) {
    writeFileSync('debug-dump.json', JSON.stringify(debugDumps, null, 2), 'utf8');
    console.log(`   （--debug：結構樣本已寫入 debug-dump.json）`);
  }
})();
