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
 *   --crawl               ⭐ 自動跟住分頁／分類連結去抓（商品分散喺多個頁面時必用）
 *   --max-pages 60        --crawl 最多抓幾多頁（預設 60）
 *   --follow-products     ⭐ 連商品詳細頁都行（網站冇商品列表頁時要用，配大 --max-pages）
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
const CRAWL        = !!flag('crawl', false);
const MAX_PAGES    = parseInt(flag('max-pages', '60'));
const FOLLOW_PRODUCTS = !!flag('follow-products', false);

if (!urls.length) {
  console.error('❌ 請提供至少一條網址，例如：\n   node tools/scrape-goods.mjs "https://chiikawapark-tokyo.jp/goods/"');
  process.exit(1);
}

// ── 優先路徑：直接讀 Next.js 嵌喺頁面嘅商品資料 ──────────────
// 好多日本商品網站用 Next.js，成個商品目錄（連分頁後面嗰啲）都已經
// 塞咗喺 self.__next_f 入面，分頁純粹係前端切開嚟顯示。
// 讀返呢份資料 = 一次過攞齊全部商品，唔使爬幾十頁，而且
// 「オンライン販売対象」呢類標示係官方欄位，比掃描文字準確得多。
function extractNextJsGoods() {
  const chunks = (self.__next_f || [])
    .map(x => (Array.isArray(x) ? x[1] : null))
    .filter(s => typeof s === 'string');
  if (!chunks.length) return null;
  const blob = chunks.join('');

  // 由 "goodsItems":[ 開始做括號配對，抽出完整 JSON 陣列
  const key = '"goodsItems":';
  const at = blob.indexOf(key);
  if (at < 0) return null;
  const start = blob.indexOf('[', at);
  if (start < 0) return null;

  let depth = 0, inStr = false, esc = false, end = -1;
  for (let i = start; i < blob.length; i++) {
    const c = blob[i];
    if (esc) { esc = false; continue; }
    if (c === '\\') { esc = true; continue; }
    if (c === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (c === '[') depth++;
    else if (c === ']') { depth--; if (depth === 0) { end = i + 1; break; } }
  }
  if (end < 0) return null;

  let items;
  try { items = JSON.parse(blob.slice(start, end)); } catch (e) { return null; }
  if (!Array.isArray(items) || !items.length) return null;

  // 價錢：「¥4,180（税込）」「各¥1,430」「単品 ¥660／ BOX ¥3,960円」
  // ⚠️ 有啲條目打錯用句號做千位分隔（例：¥1.980），要當成 1980
  const parseYen = (txt) => {
    const m = (txt || '').match(/[¥￥]\s*([0-9][0-9.,]*)|([0-9][0-9.,]*)\s*円/);
    if (!m) return 0;
    let raw = (m[1] || m[2] || '');
    raw = raw.replace(/[.,](?=\d{3}\b)/g, '');   // 千位分隔（逗號或句號）
    raw = raw.replace(/[.,]/g, '');
    const n = parseInt(raw, 10);
    return Number.isFinite(n) ? n : 0;
  };

  const origin = location.origin;
  // 圖片檔名有空格／日文／括號（例：「Gxzg_9RbsAMS591 (1).jpeg」「ふわふわ….jpg」），
  // 唔編碼就咁放入 <img src> 會載唔到，所以一律轉成合法網址。
  const abs = (u) => {
    if (!u) return '';
    const full = /^https?:/i.test(u) ? u : origin + (u.startsWith('/') ? '' : '/') + u;
    try { return encodeURI(decodeURI(full)); } catch (e) { return encodeURI(full); }
  };

  const out = [];
  for (const it of items) {
    const name = (it.title || '').replace(/\s+/g, ' ').trim();
    if (!name) continue;
    const price = parseYen(it.price);
    if (!price) continue;

    const cat02 = Array.isArray(it.category_02) ? it.category_02 : [];
    const alsoOnline = cat02.some(c => /オンライン(ストア|ショップ)?(販売)?対象(?!外)/.test(c));

    const tags = [];
    if (alsoOnline) tags.push('🛒網上都有');
    else tags.push('🏬會場限定');          // 冇標「販売対象」＝唔上網賣＝淨係現場有
    if (it.limitText) tags.push('⚠️有購買限制');

    out.push({
      name, price,
      imgUrl:   abs(it.images && it.images[0] && it.images[0].url),
      // 其餘圖片（款式／細節相）放埋入 review CSV，揀圖時唔使再開網站
      allImages: (it.images || []).map(im => abs(im && im.url)).filter(Boolean),
      link:     origin + '/goods/' + (it.id || '') + '/',
      category: it.category || '',
      tags,
      limitPerPerson: '',
      priceRaw: it.price || '',            // 原文，方便核對「単品／BOX」呢類複合價
      onlineUrl: it.purchase_url || '',
    });
  }
  return out.length ? out : null;
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
    // （略過純符號嘅「標題」，例如下拉箭嘴「▼」）
    let category = '';
    for (const h of sectionHeadings) {
      // 4 = DOCUMENT_POSITION_FOLLOWING（即 card 排喺 h 之後）
      if (h.compareDocumentPosition(card) & 4) {
        const t = clean(h.textContent);
        if (t.length >= 2 && /[\p{L}\p{N}]/u.test(t)) category = t.slice(0, 40);
      } else break;
    }

    // 標籤 + 每人限購：卡入面同所屬分類標題都掃一次
    const scanText = clean(card.textContent) + ' ' + category;
    // 去重：🏬會場限定 有兩條 pattern（対象外 / パーク限定）會夾到同一件貨
    const tags = [...new Set(TAG_PATTERNS.filter(t => t.re.test(scanText)).map(t => t.tag))];
    const lm = scanText.match(LIMIT_RE);
    const limitPerPerson = lm ? parseInt(toHalfWidth(lm[1]), 10) || '' : '';

    out.push({ name, price, imgUrl, link, category, tags, limitPerPerson });
  }

  // 3b) 單件商品詳細頁：頁面本身嗰件貨唔會排成「卡」，要另外抽返。
  //     用 og:title / h1 做名，再搵一個唔喺任何卡入面嘅價錢。
  (() => {
    const meta = n => (document.querySelector(`meta[property="${n}"]`) || {}).content || '';
    let name = clean(meta('og:title') || (document.querySelector('h1') || {}).textContent || '');
    // 剝走「｜網站名」呢類後綴
    name = name.replace(/\s*[|｜]\s*[^|｜]{0,30}$/, '').trim();
    if (!name || name.length < 3) return;
    const outside = leaves.filter(el => !cardArr.some(c => c.contains(el)));
    let price = null;
    for (const el of outside) { price = parsePrice(el.textContent); if (price) break; }
    if (!price) return;
    const scan = clean(document.body.textContent).slice(0, 4000);
    const tg = [...new Set(TAG_PATTERNS.filter(t => t.re.test(scan)).map(t => t.tag))];
    const lm = scan.match(LIMIT_RE);
    out.unshift({
      name, price,
      imgUrl: meta('og:image'), link: location.href, category: '',
      tags: tg, limitPerPerson: lm ? parseInt(toHalfWidth(lm[1]), 10) || '' : '',
    });
  })();

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
    sampleCards: Array.from(cardArr).slice(0, 3).map(c => c.outerHTML.slice(0, 900)),
    // 成頁嘅連結（睇分類導覽點樣組織，搵「全部商品」嗰類入口）
    allLinks: Array.from(document.querySelectorAll('a[href]')).slice(0, 400).map(a => ({
      href: a.getAttribute('href'),
      text: clean(a.textContent).slice(0, 40),
    })),
    // 價錢低到唔合理嘅卡，整張 HTML 拎出嚟睇下係咪認錯
    suspiciousPriceCards: products.filter(p => p.price < 200).slice(0, 3).map(p => {
      const c = cardArr.find(c => clean(c.textContent).includes(p.name.slice(0, 12)));
      return { name: p.name, price: p.price, html: c ? c.outerHTML.slice(0, 900) : '' };
    }),
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
  const header = ['商品名稱', '日幣原價', '價錢原文', '標籤', '分類', '每人限購',
                  '商品連結', '官方網購連結', '主圖', '圖片數', '全部圖片'];
  const lines = [header.map(csvCell).join(',')];
  products.forEach(p => {
    const imgs = p.allImages && p.allImages.length ? p.allImages : (p.imgUrl ? [p.imgUrl] : []);
    lines.push([
      p.name, p.price, p.priceRaw || '', (p.tags || []).join(' ／ '),
      p.category || '', p.limitPerPerson || '', p.link || '', p.onlineUrl || '',
      p.imgUrl || '', imgs.length || '', imgs.join(' '),
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

// 喺頁面度搵出「同一區段嘅其他頁」：分頁連結 + 分類頁連結
// 好多日本商品網唔係無限捲動，而係分咗好多分類頁／分頁，要逐頁去抓。
function discoverLinksInPage(baseHref) {
  const base = new URL(baseHref);
  const basePath = base.pathname.replace(/\/+$/, '');
  // ⚠️ 去重先剝尾斜線；真正去行嗰條網址一定要保留原樣，
  //    因為 /a/?page=2 同 /a?page=2 喺好多伺服器係兩件事（會 404）。
  const norm = u => u.origin + u.pathname.replace(/\/+$/, '') + u.search;
  const here = norm(base);

  const paging = new Set();
  const section = new Set();

  for (const a of document.querySelectorAll('a[href]')) {
    let u;
    try { u = new URL(a.getAttribute('href'), location.href); } catch (e) { continue; }
    if (u.origin !== base.origin) continue;

    const clean = u.origin + u.pathname + u.search;   // 保留原樣去行
    if (norm(u) === here) continue;

    // 1) 分頁：?page=2 / /page/2/ / rel=next / 文字「次へ」
    const looksPaged = /[?&](page|p|pg)=\d+/i.test(u.search) || /\/page\/\d+/i.test(u.pathname);
    const relNext = (a.getAttribute('rel') || '').toLowerCase().includes('next');
    const textNext = /次へ|次のページ|次ページ|›|»|→/.test((a.textContent || '').trim());
    if (looksPaged || relNext || textNext) { paging.add(clean); continue; }

    // 2) 同一區段嘅子頁（例如 /goods/ 下面嘅 /goods/plush/）
    if (!u.pathname.startsWith(basePath + '/')) continue;
    const tail = u.pathname.slice(basePath.length + 1).replace(/\/+$/, '');
    if (!tail) continue;
    // 排除明顯係「單件商品詳細頁」：純數字 / item-123 之類，抓佢哋好嘥時間
    if (/^\d+$/.test(tail) || /^(item|detail|product)s?[-_/]?\d+/i.test(tail)) continue;
    if (tail.split('/').length > 2) continue;   // 太深唔要
    section.add(clean);
  }

  return { paging: Array.from(paging), section: Array.from(section) };
}

async function scrapeBrowser(url, page) {
  {
    console.log('  → 開緊頁面…');
    await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });

    // 捲到底 + 撳「もっと見る」之類嘅載入更多按鈕
    // 一路捲一路報住而家有幾多件，等你睇到係咪仲載緊
    console.log('  → 捲動載入全部商品…');
    const countImgs = () => page.evaluate(() => document.querySelectorAll('img').length);
    let stable = 0, i = 0;
    for (; i < MAX_SCROLL && stable < 3; i++) {
      const before = await page.evaluate(() => document.body.scrollHeight);

      // ⚠️ 只撳「喺原地載入更多」嗰種掣（もっと見る）。
      //    「次へ」係分頁連結，撳咗會跳去第二頁，當前頁啲貨就會流失 ——
      //    分頁交畀 --crawl 逐頁去抓，所以呢度特登排除有真 href 嘅 <a>。
      const moreBtn = page
        .locator('button, [role="button"], a[href="#"], a[href^="javascript"], a:not([href])')
        .filter({ hasText: /もっと見る|もっとみる|さらに表示|もっと読み込む|Load more|See more/i })
        .first();
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

    // 先試讀 Next.js 嵌入資料：成功嘅話一次過攞齊全部商品，唔使揭頁
    const nextGoods = await page.evaluate(extractNextJsGoods).catch(() => null);
    if (nextGoods && nextGoods.length) {
      console.log(`  ⚡ 由頁面嵌入資料直接讀到 ${nextGoods.length} 件（唔使逐頁爬）`);
      return { products: nextGoods, debug: null, links: { paging: [], section: [] }, fromEmbedded: true };
    }

    console.log('  → 抽取商品資料…');
    const result = await page.evaluate(extractInPage, DEBUG);
    // 順手記低呢一頁見到嘅分頁／分類連結，畀 --crawl 用
    result.links = CRAWL ? await page.evaluate(discoverLinksInPage, url) : { paging: [], section: [] };
    return result;
  }
}

(async () => {
  const all = [];
  const debugDumps = [];

  // 開一次瀏覽器，全部頁面共用（快好多）
  let browser = null, page = null;
  if (!STATIC) {
    const { chromium } = await import('playwright').catch(() => {
      throw new Error('搵唔到 playwright，請先行：npm install playwright && npx playwright install chromium');
    });
    // --chrome-path：如果 npx playwright install 裝唔到，可以指定現成嘅 Chrome/Chromium
    browser = await chromium.launch(CHROME_PATH ? { executablePath: CHROME_PATH } : {});
    // 扮返個正常日本訪客：預設 headless UA 會寫住 HeadlessChrome，有啲網站會擋
    page = await browser.newPage({
      viewport:   { width: 1440, height: 900 },
      userAgent:  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
      locale:     'ja-JP',
      timezoneId: 'Asia/Tokyo',
      extraHTTPHeaders: { 'Accept-Language': 'ja-JP,ja;q=0.9' },
    });
  }

  // 去重用嘅正規化：剝走尾斜線做比較，但行嘅時候用原本條網址
  const normKey = (u) => {
    try { const x = new URL(u); return x.origin + x.pathname.replace(/\/+$/, '') + x.search; }
    catch { return u.replace(/\/+$/, ''); }
  };

  const queue   = [...urls];
  const visited = new Set();
  // 商品卡指向嘅網址 = 單件商品詳細頁。爬佢哋好嘥時間（每頁得幾件「相關商品」），
  // 而且會令人以為抓齊咗。呢個判斷唔使靠網址格式，適用於任何網站。
  const detailPages = new Set();
  const startKeys = new Set(urls.map(normKey));
  let pagesDone = 0, skippedDetail = 0;

  try {
    while (queue.length && pagesDone < MAX_PAGES) {
      const url = queue.shift();
      const key = normKey(url);
      if (visited.has(key)) continue;
      // 商品詳細頁：預設略過（爬列表頁效率高好多）。
      // 但有啲網站根本冇商品列表頁，啲貨淨係靠「相關商品」互相連住 ——
      // 嗰陣就要 --follow-products，行勻成個商品網絡。
      if (detailPages.has(key) && !startKeys.has(key) && !FOLLOW_PRODUCTS) { skippedDetail++; continue; }
      visited.add(key);
      pagesDone++;

      console.log(`\n📦 [${pagesDone}/${Math.min(MAX_PAGES, visited.size + queue.length)}] ${url}`);
      try {
        const r = STATIC ? await scrapeStatic(url) : await scrapeBrowser(url, page);
        console.log(`  ✅ 抓到 ${r.products.length} 件商品`);
        all.push(...r.products);
        if (r.debug) debugDumps.push({ url, ...r.debug });

        // 由嵌入資料一次過攞齊咗，就唔使再爬任何子頁
        if (r.fromEmbedded) { queue.length = 0; }

        // 商品卡指向嘅網址記低做「詳細頁」，之後唔會再爬
        for (const p of r.products) if (p.link) detailPages.add(normKey(p.link));

        // --crawl：自動排隊去抓分頁同分類頁
        if (CRAWL && r.links) {
          const queued = new Set(queue.map(normKey));
          const cand = [...r.links.paging, ...r.links.section];
          // --follow-products：連商品詳細頁都排隊（用嚟行商品網絡）
          if (FOLLOW_PRODUCTS) for (const p of r.products) if (p.link) cand.push(p.link);
          const fresh = cand.filter(u => {
            const k = normKey(u);
            return !visited.has(k) && !queued.has(k) &&
                   (FOLLOW_PRODUCTS || !detailPages.has(k));
          });
          if (fresh.length) {
            queue.push(...fresh);
            console.log(`  🔗 發現 ${fresh.length} 條新子頁（分頁 ${r.links.paging.length}／分類 ${r.links.section.length}）`);
          }
        }
      } catch (e) {
        console.error(`  ❌ 失敗：${e.message}`);
      }
    }

    if (skippedDetail) {
      console.log(`\n  ⏭️  略過咗 ${skippedDetail} 條單件商品詳細頁（唔係商品列表，爬佢哋冇用）`);
    }
    if (queue.length) {
      console.warn(`\n  ⚠️  仲有 ${queue.length} 條子頁未抓（已達上限 --max-pages ${MAX_PAGES}）`);
      console.warn(`      加大：--max-pages ${MAX_PAGES * 2}`);
    }
  } finally {
    if (browser) await browser.close();
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
      console.warn(`         唔好就咁攞去報價。按呢個次序試：`);
      if (!CRAWL) {
        console.warn(`         1. ⭐ 加 --crawl —— 商品好大機會分咗喺多個分類頁／分頁，`);
        console.warn(`               呢個選項會自動跟住連結逐頁抓。多數係呢個原因。`);
      } else {
        console.warn(`         1. 加大頁數上限：--max-pages ${MAX_PAGES * 2}`);
      }
      console.warn(`         2. 加大捲動次數：--max-scroll ${MAX_SCROLL * 2}`);
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
