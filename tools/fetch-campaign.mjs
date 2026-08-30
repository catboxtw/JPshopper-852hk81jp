#!/usr/bin/env node
/**
 * 抓 Nissen ／ ZOZOTOWN 嘅推介同特價商品，整成一版「出 post 揀相」頁。
 *
 *   node tools/fetch-campaign.mjs                       # 兩間鋪都抓
 *   node tools/fetch-campaign.mjs --site zozo --limit 30
 *   node tools/fetch-campaign.mjs --url https://zozo.jp/sale/ --site zozo
 *
 * 出三樣嘢入 data/：
 *   campaign.html          ── 圖片牆，喺手機揀相出 post 用
 *   campaign-<日期>.csv    ── 同一批貨嘅表格，開 Excel 睇
 *   campaign.json          ── 機器讀嘅版本（下次抓可以比對出新貨）
 *
 * 抓唔到嘢嘅時候加 --debug，會出 campaign-debug.json，
 * 入面有頁面真實結構，唔使估。
 */

import { writeFileSync, mkdirSync, existsSync, readFileSync } from 'fs';
import { join } from 'path';

// ── 指令參數 ──────────────────────────────────────────────
const argv = process.argv.slice(2);
const flag = (name, def = null) => {
  const i = argv.indexOf(`--${name}`);
  if (i === -1) return def;
  const next = argv[i + 1];
  return (!next || next.startsWith('--')) ? true : next;
};
const flagAll = (name) => argv.reduce((acc, a, i) => {
  if (a === `--${name}` && argv[i + 1] && !argv[i + 1].startsWith('--')) acc.push(argv[i + 1]);
  return acc;
}, []);

const OUT_DIR   = flag('out', 'data');
const LIMIT     = parseInt(flag('limit', '24')) || 24;
const DEBUG     = !!flag('debug', false);
const HEADFUL   = !!flag('headful', false);
const TIMEOUT   = parseInt(flag('timeout', '45000')) || 45000;
const MAX_SCROLL = parseInt(flag('max-scroll', '12')) || 12;

const SUPA_URL = flag('supabase-url', 'https://pksqfpirggvsftvqrtji.supabase.co');
const SUPA_KEY = flag('supabase-key', process.env.SUPABASE_ANON_KEY || '');
let RATE_HK = parseFloat(flag('rate-hk', '0.057'));
let RATE_TW = parseFloat(flag('rate-tw', '0.24'));

// ── 每間鋪抓邊幾版 ────────────────────────────────────────
const SITES = {
  zozo: {
    label: 'ZOZOTOWN',
    host: 'zozo.jp',
    sources: [
      { name: '人氣排行', url: 'https://zozo.jp/ranking/' },
      { name: '特價',     url: 'https://zozo.jp/sale/' },
    ],
  },
  nissen: {
    label: 'Nissen',
    host: 'nissen.co.jp',
    sources: [
      { name: '人氣排行', url: 'https://www.nissen.co.jp/ranking/' },
      { name: '特價',     url: 'https://www.nissen.co.jp/sale/' },
      { name: '首頁推介', url: 'https://www.nissen.co.jp/' },
    ],
  },
};

const wantedSites = flagAll('site');
const customUrls  = flagAll('url');
const siteKeys = wantedSites.length ? wantedSites : Object.keys(SITES);

// ── 共用小工具（全部要喺瀏覽器入面行，所以寫成字串注入）──
// 呢段會 serialize 落 page.evaluate，唔可以引用外面嘅變數。
function extractInPage() {
  const origin = location.origin;
  const out = [];
  const seen = new Set();

  // 日圓可以寫成 ¥1,980 或者 ¥1.980（有啲版用句點做千位分隔）
  const parseYen = (txt) => {
    if (!txt) return 0;
    const m = String(txt).match(/[¥￥]\s*([0-9][0-9.,]*)|([0-9][0-9.,]*)\s*円/);
    if (!m) return 0;
    let raw = m[1] || m[2] || '';
    raw = raw.replace(/[.,](?=\d{3}\b)/g, '').replace(/[.,]/g, '');
    return parseInt(raw, 10) || 0;
  };

  // 圖片網址可能有空格／日文／★，唔編碼就會靜靜地載入失敗
  const abs = (u) => {
    if (!u) return '';
    let full = /^https?:/i.test(u) ? u : (u.startsWith('//') ? location.protocol + u
              : origin + (u.startsWith('/') ? '' : '/') + u);
    try { return encodeURI(decodeURI(full)); } catch (e) { return encodeURI(full); }
  };

  const push = (it) => {
    if (!it || !it.image || !it.url) return;
    const key = it.url.split('?')[0];
    if (seen.has(key)) return;
    seen.add(key);
    out.push(it);
  };

  // ── 策略 1：__NEXT_DATA__ ／ __next_f 入面嘅商品物件 ──
  // 唔寫死 props 路徑 —— 排行版同商品版嘅路徑唔一樣，
  // 所以行勻成棵樹，見到「似商品」嘅物件就收。
  const collectFromJson = (root) => {
    const stack = [root];
    let guard = 0;
    while (stack.length && guard++ < 200000) {
      const node = stack.pop();
      if (!node || typeof node !== 'object') continue;
      if (Array.isArray(node)) { for (const v of node) stack.push(v); continue; }

      const name = node.goodsName || node.itemName || node.productName || node.name;
      const img  = node.defaultImageUrl || node.imageUrl || node.image ||
                   (node.images && node.images[0] &&
                     (node.images[0].url || node.images[0].imageUrl || node.images[0]));
      const pi   = node.priceInfo || node.price || {};
      const price = typeof pi === 'object'
        ? parseInt(pi.price || pi.salePrice || pi.value || 0)
        : parseYen(pi) || parseInt(pi) || 0;

      if (name && img && price > 0) {
        const href = node.goodsUrl || node.url || node.link || node.detailUrl || '';
        push({
          name: String(name).trim(),
          image: abs(typeof img === 'string' ? img : ''),
          url: abs(href),
          yen: price,
          origYen: (pi && pi.doublePriceLabel && parseInt(pi.doublePriceLabel.price)) ||
                   parseInt(node.listPrice || 0) || 0,
          discount: (pi && pi.discountRate) || node.discountRate || 0,
          brand: (node.brand && (node.brand.brandName || node.brand.name)) ||
                 node.brandName || '',
          from: 'embedded',
        });
      }
      for (const k in node) {
        const v = node[k];
        if (v && typeof v === 'object') stack.push(v);
      }
    }
  };

  const nd = document.getElementById('__NEXT_DATA__');
  if (nd) { try { collectFromJson(JSON.parse(nd.textContent)); } catch (e) {} }

  if (out.length === 0 && Array.isArray(self.__next_f)) {
    const blob = self.__next_f.map(x => Array.isArray(x) ? x[1] : null)
                              .filter(s => typeof s === 'string').join('');
    // 揀出最大嗰個 JSON 陣列／物件試吓解
    for (const key of ['"goodsItems":', '"items":', '"products":', '"ranking":']) {
      const at = blob.indexOf(key);
      if (at < 0) continue;
      const start = blob.indexOf('[', at);
      if (start < 0) continue;
      let depth = 0, inStr = false, esc = false, end = -1;
      for (let i = start; i < blob.length; i++) {
        const c = blob[i];
        if (esc) { esc = false; continue; }
        if (c === '\\') { esc = true; continue; }
        if (c === '"') { inStr = !inStr; continue; }
        if (inStr) continue;
        if (c === '[') depth++;
        else if (c === ']') { depth--; if (depth === 0) { end = i; break; } }
      }
      if (end > start) {
        try { collectFromJson(JSON.parse(blob.slice(start, end + 1))); } catch (e) {}
      }
      if (out.length) break;
    }
  }

  // ── 策略 2：JSON-LD ──
  if (out.length === 0) {
    document.querySelectorAll('script[type="application/ld+json"]').forEach(s => {
      try { collectFromJson(JSON.parse(s.textContent)); } catch (e) {}
    });
  }

  // ── 策略 3：老實睇 DOM ──
  // 一個商品卡＝一條 <a>，入面有圖，subtree 有個日圓價。
  if (out.length === 0) {
    document.querySelectorAll('a[href]').forEach(a => {
      const href = a.getAttribute('href') || '';
      if (!href || href.startsWith('javascript') || href === '#') return;

      const im = a.querySelector('img');
      if (!im) return;
      const src = im.getAttribute('src') || im.getAttribute('data-src') ||
                  im.getAttribute('data-original') || '';
      if (!src) return;
      if (/\/icon|\/logo|\/common\/|\/btn|\/arrow|\/parts\/|sprite|blank\.|spacer/i.test(src)) return;

      // 價錢：睇卡入面每個元素嘅「直屬文字」，
      // 因為 <p>¥2,530<span>税込</span></p> 呢種只睇葉節點會走漏。
      let yen = 0;
      const walk = a.querySelectorAll('*');
      for (const el of [a, ...walk]) {
        let direct = '';
        for (const n of el.childNodes) if (n.nodeType === 3) direct += n.nodeValue;
        const v = parseYen(direct);
        if (v > 0) { yen = yen ? Math.min(yen, v) : v; }
      }
      if (yen <= 0) return;

      const alt = (im.getAttribute('alt') || '').trim();
      const txt = (a.innerText || '').trim().split('\n').map(s => s.trim())
                    .filter(s => s && !/[¥￥]|円|%|OFF/i.test(s))[0] || '';
      push({
        name: (alt || txt).slice(0, 120),
        image: abs(src),
        url: abs(href),
        yen,
        origYen: 0,
        discount: 0,
        brand: '',
        from: 'dom',
      });
    });
  }

  return {
    items: out,
    diag: {
      url: location.href,
      title: document.title,
      hasNextData: !!nd,
      hasNextF: Array.isArray(self.__next_f),
      anchorCount: document.querySelectorAll('a[href]').length,
      imgCount: document.querySelectorAll('img').length,
      // 抓唔到嘢嘅時候，呢段係最有用嘅線索
      sampleClasses: [...document.querySelectorAll('li,article,div')]
        .slice(0, 400)
        .map(e => e.className)
        .filter(c => typeof c === 'string' && c && /item|goods|product|card|rank/i.test(c))
        .slice(0, 30),
    },
  };
}

// ── 匯率：同 admin 嗰版設定一致 ────────────────────────────
async function loadRates() {
  if (!SUPA_KEY) return;
  try {
    const res = await fetch(`${SUPA_URL}/rest/v1/app_settings?select=key,value`, {
      headers: { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}` },
    });
    if (!res.ok) return;
    const rows = await res.json();
    for (const r of rows) {
      if (r.key === 'rate_hk' && parseFloat(r.value) > 0) RATE_HK = parseFloat(r.value);
      if (r.key === 'rate_tw' && parseFloat(r.value) > 0) RATE_TW = parseFloat(r.value);
    }
    console.log(`  匯率跟 admin 設定：HK ${RATE_HK} ／ TW ${RATE_TW}`);
  } catch (e) {
    console.log(`  ⚠️  讀唔到匯率設定，用預設值（${e.message}）`);
  }
}

// ── CSV ───────────────────────────────────────────────────
const csvCell = v => {
  const s = v === null || v === undefined ? '' : String(v);
  return /[",\r\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
};
function toCsv(items) {
  const head = ['店舖', '來源', '商品名', '日圓', '原價', '折扣%', '港幣', '台幣', '品牌', '圖片', '商品網址'];
  const rows = items.map(i => [
    i.site, i.source, i.name, i.yen, i.origYen || '', i.discount || '',
    Math.ceil(i.yen * RATE_HK), Math.ceil(i.yen * RATE_TW),
    i.brand || '', i.image, i.url,
  ]);
  // BOM 令 Excel 認得 UTF-8，CRLF 令佢唔會撈埋一行
  return '﻿' + [head, ...rows].map(r => r.map(csvCell).join(',')).join('\r\n') + '\r\n';
}

// ── 圖片牆 ────────────────────────────────────────────────
function toHtml(items, meta) {
  const card = (i, n) => `
    <figure class="card">
      <a class="shot" href="${i.image}" target="_blank" rel="noopener">
        <img src="${i.image}" alt="" loading="lazy">
        ${i.discount ? `<span class="off">${i.discount}%OFF</span>` : ''}
        <span class="no">${n}</span>
      </a>
      <figcaption>
        ${i.brand ? `<p class="brand">${esc(i.brand)}</p>` : ''}
        <p class="name">${esc(i.name)}</p>
        <p class="price">
          <span class="jpy">¥${i.yen.toLocaleString()}</span>
          ${i.origYen && i.origYen > i.yen ? `<s>¥${i.origYen.toLocaleString()}</s>` : ''}
        </p>
        <p class="local">
          <span>HK$${Math.ceil(i.yen * RATE_HK).toLocaleString()}</span>
          <span>NT$${Math.ceil(i.yen * RATE_TW).toLocaleString()}</span>
        </p>
        <a class="src" href="${i.url}" target="_blank" rel="noopener">${esc(i.site)} · ${esc(i.source)} ↗</a>
      </figcaption>
    </figure>`;

  return `<!doctype html>
<html lang="zh-HK"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>今期推介 · 出 post 揀相</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+HK:wght@400;500;700&family=IBM+Plex+Mono:wght@500&display=swap">
<style>
  :root{
    --paper:#f4f2ed; --ink:#17161a; --muted:#77736c; --line:#e2ded6;
    --card:#fffdf9; --hk:#c0392b; --tw:#2980b9;
  }
  @media (prefers-color-scheme:dark){:root:not([data-theme="light"]){
    --paper:#151417; --ink:#f0ede7; --muted:#96918a; --line:#2b292d; --card:#1e1d21;
    --hk:#e35d4c; --tw:#5aa5db;
  }}
  :root[data-theme="dark"]{
    --paper:#151417; --ink:#f0ede7; --muted:#96918a; --line:#2b292d; --card:#1e1d21;
    --hk:#e35d4c; --tw:#5aa5db;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--paper);color:var(--ink);
       font-family:"Noto Sans HK",system-ui,-apple-system,sans-serif;
       -webkit-font-smoothing:antialiased}
  header{padding:28px 20px 18px;border-bottom:1px solid var(--line)}
  h1{margin:0 0 6px;font-size:20px;font-weight:700;letter-spacing:.01em}
  .meta{margin:0;font-size:12px;color:var(--muted);
        font-family:"IBM Plex Mono",ui-monospace,monospace}
  .tips{margin:14px 20px 0;padding:12px 14px;border:1px dashed var(--line);
        border-radius:8px;font-size:12px;line-height:1.7;color:var(--muted)}
  .grid{display:grid;gap:14px;padding:18px 20px 60px;
        grid-template-columns:repeat(auto-fill,minmax(168px,1fr))}
  .card{margin:0;background:var(--card);border:1px solid var(--line);
        border-radius:10px;overflow:hidden;display:flex;flex-direction:column}
  .shot{position:relative;display:block;aspect-ratio:1;background:var(--paper)}
  .shot img{width:100%;height:100%;object-fit:cover;display:block}
  .off{position:absolute;left:8px;top:8px;background:var(--hk);color:#fff;
       font-size:10px;font-weight:700;padding:3px 6px;border-radius:4px;
       font-family:"IBM Plex Mono",monospace}
  .no{position:absolute;right:8px;top:8px;background:rgba(23,22,26,.72);color:#fff;
      font-size:10px;padding:3px 6px;border-radius:4px;
      font-family:"IBM Plex Mono",monospace}
  figcaption{padding:10px 11px 12px;display:flex;flex-direction:column;gap:4px;flex:1}
  .brand{margin:0;font-size:10px;letter-spacing:.06em;text-transform:uppercase;color:var(--muted)}
  .name{margin:0;font-size:12px;line-height:1.5;font-weight:500;
        display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical;overflow:hidden}
  .price{margin:2px 0 0;font-size:13px;font-weight:700;
         font-family:"IBM Plex Mono",monospace;font-variant-numeric:tabular-nums}
  .price s{font-size:11px;font-weight:500;color:var(--muted);margin-left:6px}
  .local{margin:0;display:flex;gap:10px;font-size:11px;color:var(--muted);
         font-family:"IBM Plex Mono",monospace;font-variant-numeric:tabular-nums}
  .local span:first-child{color:var(--hk)} .local span:last-child{color:var(--tw)}
  .src{margin-top:auto;padding-top:6px;font-size:10px;color:var(--muted);text-decoration:none}
  .src:hover{color:var(--ink)}
  a:focus-visible{outline:2px solid var(--tw);outline-offset:2px}
  .empty{padding:60px 20px;text-align:center;color:var(--muted);font-size:13px;line-height:1.8}
</style></head><body>
<header>
  <h1>今期推介 · 出 post 揀相</h1>
  <p class="meta">${meta.date}　·　${items.length} 件　·　¥1 = HK$${RATE_HK} / NT$${RATE_TW}</p>
</header>
<p class="tips">
  撳張相會開原圖 —— 喺手機長撳就儲存得，直接攞去出 post。<br>
  撳最底行字會去返商品原頁，落單／查尺碼用。
</p>
${items.length ? `<div class="grid">${items.map((i, n) => card(i, n + 1)).join('')}</div>`
  : `<p class="empty">今次抓唔到貨。<br>去 Actions 嗰個 run 下載 <code>campaign-debug.json</code> 睇下頁面結構變咗啲乜。</p>`}
</body></html>`;
}

const esc = s => String(s || '').replace(/[&<>"]/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

// ── 主流程 ────────────────────────────────────────────────
async function main() {
  const { chromium } = await import('playwright');

  await loadRates();

  const jobs = [];
  if (customUrls.length) {
    const key = siteKeys[0] || 'zozo';
    customUrls.forEach(u => jobs.push({ site: SITES[key]?.label || key, source: '指定網址', url: u, key }));
  } else {
    for (const key of siteKeys) {
      const s = SITES[key];
      if (!s) { console.log(`⚠️  唔識「${key}」呢間鋪，跳過`); continue; }
      s.sources.forEach(src => jobs.push({ site: s.label, source: src.name, url: src.url, key }));
    }
  }

  const browser = await chromium.launch({
    headless: !HEADFUL,
    executablePath: process.env.CHROME_PATH || undefined,
  });
  const ctx = await browser.newContext({
    locale: 'ja-JP',
    userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 ' +
               '(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
    viewport: { width: 1280, height: 1600 },
  });

  const all = [];
  const diags = [];

  for (const job of jobs) {
    console.log(`\n▶ ${job.site} · ${job.source}\n  ${job.url}`);
    const page = await ctx.newPage();
    try {
      const resp = await page.goto(job.url, { waitUntil: 'domcontentloaded', timeout: TIMEOUT });
      const status = resp ? resp.status() : 0;
      if (status >= 400) {
        console.log(`  ❌ HTTP ${status}，跳過`);
        diags.push({ ...job, status, error: `HTTP ${status}` });
        await page.close();
        continue;
      }

      // 圖片多數係 lazy load，捲落去先影得到
      for (let i = 0; i < MAX_SCROLL; i++) {
        await page.evaluate(() => window.scrollBy(0, window.innerHeight * 0.9));
        await page.waitForTimeout(350);
      }
      await page.waitForTimeout(800);

      const r = await page.evaluate(extractInPage);
      diags.push({ ...job, status, ...r.diag, found: r.items.length });

      const tagged = r.items
        .filter(i => i.yen > 0 && i.image)
        .map(i => ({ ...i, site: job.site, source: job.source }));
      console.log(`  ✅ ${tagged.length} 件（${r.items[0]?.from || '冇'}）`);
      all.push(...tagged);
    } catch (e) {
      console.log(`  ❌ ${e.message}`);
      diags.push({ ...job, error: e.message });
    }
    await page.close();
  }

  await browser.close();

  // 同一件貨可能喺排行同特價都出現，夾埋去重（留折扣大嗰個）
  const byUrl = new Map();
  for (const it of all) {
    const k = (it.url || it.image).split('?')[0];
    const prev = byUrl.get(k);
    if (!prev || (it.discount || 0) > (prev.discount || 0)) byUrl.set(k, it);
  }
  // 有折扣嘅行先，出 post 最好賣
  const items = [...byUrl.values()]
    .sort((a, b) => (b.discount || 0) - (a.discount || 0) || a.yen - b.yen)
    .slice(0, LIMIT);

  if (!existsSync(OUT_DIR)) mkdirSync(OUT_DIR, { recursive: true });
  const date = new Date().toISOString().slice(0, 10);

  writeFileSync(join(OUT_DIR, 'campaign.html'), toHtml(items, { date }));
  writeFileSync(join(OUT_DIR, `campaign-${date}.csv`), toCsv(items));
  writeFileSync(join(OUT_DIR, 'campaign.json'),
    JSON.stringify({ date, rateHk: RATE_HK, rateTw: RATE_TW, items }, null, 2));

  if (DEBUG) {
    writeFileSync(join(OUT_DIR, 'campaign-debug.json'), JSON.stringify(diags, null, 2));
  }

  console.log(`\n──────────────────────────────`);
  console.log(`共 ${items.length} 件（原始 ${all.length} 件，去重後揀頭 ${LIMIT}）`);
  console.log(`campaign.html ／ campaign-${date}.csv ／ campaign.json → ${OUT_DIR}/`);
  if (items.length === 0) {
    console.log(`\n⚠️  一件都抓唔到。加 --debug 再行，睇 campaign-debug.json 入面嘅頁面結構。`);
    process.exitCode = 1;
  }
}

main().catch(e => { console.error(e); process.exit(1); });
