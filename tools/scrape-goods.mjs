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
 *   --out goods.csv       輸出檔名（預設 goods.csv）
 *   --rate 0.22           日幣→台幣匯率（預設 0.22）
 *   --debug               抓唔到嘢時用：印出頁面結構樣本，send 俾 Claude 睇
 *   --static              唔用瀏覽器，直接抓 HTML（快，但 JS 載入嘅頁面會抓唔到）
 *   --max-scroll 40       最多向下捲幾多次（處理無限捲動，預設 30）
 *
 * 輸出 CSV 欄位順序 = Google Sheet 商品分頁欄位順序（A~J），
 * 開啟後由 A2 開始貼上即可。重量（G欄）留空，要你自己填，
 * 填完 D/E/I 欄嘅公式會自動計出成本同售價。
 */

import { writeFileSync } from 'node:fs';

// ── 解析指令參數 ──────────────────────────────────────────
const argv = process.argv.slice(2);
const urls = argv.filter(a => /^https?:\/\//.test(a));
const flag = (name, def = null) => {
  const i = argv.indexOf('--' + name);
  return i >= 0 ? (argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[i + 1] : true) : def;
};
const OUT        = flag('out', 'goods.csv');
const TWD_RATE   = parseFloat(flag('rate', '0.22'));
const DEBUG      = !!flag('debug', false);
const STATIC     = !!flag('static', false);
const MAX_SCROLL = parseInt(flag('max-scroll', '30'));

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

    out.push({ name, price, imgUrl, link, category });
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
      p.category ? '分類：' + p.category : '', // H 備註
      `=CEILING(C${r}*${rate},5)`,          // I 台幣售價（進位至5的倍數）
      '',                                   // J 限購：留空 = 不限
    ].map(csvCell).join(','));
  });

  return '﻿' + lines.join('\r\n'); // BOM：確保 Excel 開日文唔會亂碼
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

  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  try {
    console.log('  → 開緊頁面…');
    await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });

    // 捲到底 + 撳「もっと見る」之類嘅載入更多按鈕
    console.log('  → 捲動載入全部商品…');
    let stable = 0;
    for (let i = 0; i < MAX_SCROLL && stable < 3; i++) {
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
  const products = all.filter(p => {
    const k = p.name + '|' + p.price;
    if (seen.has(k)) return false;
    seen.add(k);
    return true;
  });

  if (!products.length) {
    console.error('\n❌ 一件商品都抓唔到。');
    console.error('   請加 --debug 再行一次，然後將印出嚟嘅內容 send 俾 Claude 幫你調整。');
    if (debugDumps.length) {
      writeFileSync('debug-dump.json', JSON.stringify(debugDumps, null, 2), 'utf8');
      console.error('   （已寫入 debug-dump.json）');
    }
    process.exit(1);
  }

  writeFileSync(OUT, toCsv(products, TWD_RATE), 'utf8');
  console.log(`\n🎉 完成！共 ${products.length} 件商品 → ${OUT}`);
  console.log(`   價錢範圍：¥${Math.min(...products.map(p => p.price)).toLocaleString()} ~ ¥${Math.max(...products.map(p => p.price)).toLocaleString()}`);
  console.log(`\n   下一步：用 Excel／Google Sheet 開 ${OUT}，`);
  console.log(`   填好「重量g」欄之後，成本／港幣售價／台幣售價會自動計出嚟。`);

  if (debugDumps.length) {
    writeFileSync('debug-dump.json', JSON.stringify(debugDumps, null, 2), 'utf8');
    console.log(`   （--debug：結構樣本已寫入 debug-dump.json）`);
  }
})();
