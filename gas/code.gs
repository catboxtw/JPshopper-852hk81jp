// =================================================================
// onOpen：Google Sheet 頂部自定義同步選單
// =================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🗂 852hk.81jp 同步")
    .addItem("同步全部活動商品", "menuSyncAllEvents")
    .addItem("同步易寄取地址", "menuSyncAddresses")
    .addSeparator()
    .addItem("同步全部（活動 + 地址）", "menuSyncAll")
    .addSeparator()
    .addItem("手動同步已購數量", "menuSyncOrderedQty")
    .addToUi();
}

function menuSyncAllEvents() {
  var res = syncAllActiveEventsToSupabase();
  var detail = res.details ? res.details.map(function(d) {
    return d.eventName + "：" + d.result + (d.synced ? " (" + d.synced + "件)" : "") + (d.message ? " ⚠" + d.message : "");
  }).join("\n") : "";
  SpreadsheetApp.getUi().alert("活動同步完成",
    "成功：" + res.ok + " 個 ／ 失敗：" + res.failed + " 個\n\n" + detail,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuSyncAddresses() {
  syncHKAddressesToSupabase();
  SpreadsheetApp.getUi().alert("易寄取地址同步完成", "地址已成功同步到 Supabase。", SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuSyncAll() {
  var evRes = syncAllActiveEventsToSupabase();
  syncHKAddressesToSupabase();
  SpreadsheetApp.getUi().alert("全部同步完成",
    "【活動】成功 " + evRes.ok + " 個，失敗 " + evRes.failed + " 個\n【地址】同步完成",
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuSyncOrderedQty() {
  syncOrderedQtyToSupabase();
  SpreadsheetApp.getUi().alert("完成", "已購數量同步完成。", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==================== 🛠️ 設定區 ====================
var ADMIN_PAGE_URL = "https://jpshopper-852hk81jp.vercel.app/admin.html";
var MY_NOTIFICATION_EMAIL = "852hk81jp@gmail.com";
var LINE_URL = "https://line.me/ti/p/tf98MWHnC5";
// ==================================================

// =================================================================
// Sheet Helper 函式
// =================================================================
// 舊版總表（向下相容）
function getPurchaseSheet(ss) {
  return ss.getSheetByName("[Data]購貨紀錄") ||
         ss.getSheetByName("購貨紀錄") || null;
}
function getProfitSheet(ss) {
  return ss.getSheetByName("[Data]盈利紀錄") ||
         ss.getSheetByName("盈利紀錄") || null;
}

// 模糊比對商品名取圖（頂層函式，供所有 email 函式使用）
function findPhotoByName(photoMap, productName) {
  if (!productName) return "";
  if (photoMap[productName]) return photoMap[productName];
  var cleanTarget = productName.toLowerCase().replace(/[\s　]/g, "");
  for (var pk in photoMap) {
    var cleanKey = pk.toLowerCase().replace(/[\s　]/g, "");
    if (cleanTarget === cleanKey || cleanTarget.indexOf(cleanKey) !== -1 || cleanKey.indexOf(cleanTarget) !== -1) {
      return photoMap[pk];
    }
  }
  return "";
}

// 去除 [END] 前綴（用於查找 Data sheet，sheet 名稱不變）
function stripEndPrefix(name) {
  if (!name) return name;
  return name.replace(/^\[END\]\s*/i, '');
}

// Event product sheet 查找（三種方式）
// 情境1: eventName="[END]0525 Popup" → 直接找 "[END]0525 Popup"
// 情境2: eventName="0525 Popup"（訂單紀錄C欄的舊名）→ 找 "[END]0525 Popup"
// 情境3: eventName="0614 Event"（現有活動）→ 直接找 "0614 Event"
function getSheetByEventName(ss, eventName) {
  if (!eventName) return null;
  var s;
  // 嘗試1：精確比對（含或不含[END]）
  s = ss.getSheetByName(eventName);
  if (s) return s;
  // 嘗試2：去掉[END]前綴
  var clean = stripEndPrefix(eventName);
  if (clean !== eventName) {
    s = ss.getSheetByName(clean);
    if (s) return s;
  }
  // 嘗試3：加上[END]前綴（訂單記錄了舊名，但sheet已改名為[END]開頭）
  var withEnd = "[END]" + clean;
  if (withEnd !== eventName) {
    s = ss.getSheetByName(withEnd);
    if (s) return s;
  }
  return null;
}

// 新版 per-event sheet：[Data](eventName)購貨紀錄
// 欄位：A=商品名稱, B=款式, C=日幣, D=已付款, E=未付款, F=已購買, G=已結算, H+=批次日期
function getEventPurchaseSheet(ss, eventName, createIfMissing) {
  var sheetName = "[Data](" + stripEndPrefix(eventName) + ")購貨紀錄";
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet && createIfMissing) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["商品名稱", "款式", "日幣單價", "已付款數量", "未付款數量", "已購買數量", "已結算數量"]);
  }
  return sheet;
}

function hasEventPurchaseSheet(ss, eventName) {
  return !!ss.getSheetByName("[Data](" + stripEndPrefix(eventName) + ")購貨紀錄");
}

function jsonpOrJson_(param, data) {
  var callback = param.callback || "";

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(data) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var param = (e && e.parameter) ? e.parameter : {};
  var action = param.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === "syncEventToSupabase") {
    var evName = param.eventName || param.event || "";
    var syncRes;

    try {
      if (!evName) {
        syncRes = { result: "error", message: "缺少 eventName" };
      } else {
        syncRes = syncEventToSupabase(ss, evName);
      }
    } catch (err) {
      syncRes = {
        result: "error",
        message: err && err.stack ? err.stack : String(err)
      };
    }

    return jsonpOrJson_(param, syncRes);
  }

  if (action === "syncAllEventsToSupabase") {
    var syncAllRes;

    try {
      syncAllRes = syncAllActiveEventsToSupabase();
    } catch (err) {
      syncAllRes = {
        result: "error",
        total: 0,
        ok: 0,
        failed: 1,
        message: err && err.stack ? err.stack : String(err)
      };
    }

    return jsonpOrJson_(param, syncAllRes);
  }

  // 🏛️ 【分流 1】大廳首頁：回傳所有分頁名稱
  if (action === "getSheets") {
    var sheets = ss.getSheets();
    var sheetNames = [];
    var sysSheets = ["訂單紀錄", "易寄取地址", "Blank", "收單截止時間", "購貨紀錄", "盈利紀錄"];
    for (var i = 0; i < sheets.length; i++) {
      var sn = sheets[i].getName();
      // 排除 [Data] 前綴的 system sheet 和固定名稱
      if (sn.startsWith("[Data]")) continue;
      if (sysSheets.indexOf(sn) !== -1) continue;
      sheetNames.push(sn);
    }
    return ContentService.createTextOutput(JSON.stringify(sheetNames))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  // 📋 【分流 2】後台管理：撈取所有訂單紀錄
  if (action === "getOrders" || action === "getAdminOrders") {
    var sheet = ss.getSheetByName("訂單紀錄");
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({orders: []})).setMimeType(ContentService.MimeType.JSON);
    
    var rows = sheet.getDataRange().getValues();
    var orders = [];

    // 預先建立所有 event 分頁的商品圖片對照表
    // { eventName: { "商品名": photoUrl } }
    var eventPhotoCache = {};
    function getEventPhotoMap(evName) {
      if (!evName) return {};
      if (eventPhotoCache[evName]) return eventPhotoCache[evName];
      var evSheet = getSheetByEventName(ss, evName);
      var map = {};
      if (evSheet) {
        var evRows = evSheet.getDataRange().getValues();
        for (var ep = 1; ep < evRows.length; ep++) {
          var pn = evRows[ep][0] ? evRows[ep][0].toString().trim() : "";
          var ph = evRows[ep][5] ? evRows[ep][5].toString().trim() : "";
          if (pn) map[pn] = ph;
        }
      }
      eventPhotoCache[evName] = map;
      return map;
    }

    // 解析 summary 取出每個商品名稱
    function extractProductNames(summaryText) {
      // 直接用 parseSummaryItems 確保與解析邏輯一致
      var parsed = parseSummaryItems(summaryText);
      var names = [];
      for (var pi = 0; pi < parsed.length; pi++) {
        var n = parsed[pi].name;
        if (n && names.indexOf(n) === -1) names.push(n);
      }
      return names;
    }

    // 模糊比對商品名取圖
    function findPhoto(photoMap, productName) {
      if (!productName) return "";
      if (photoMap[productName]) return photoMap[productName];
      var cleanTarget = productName.toLowerCase().replace(/[\s　]/g, "");
      for (var pk in photoMap) {
        var cleanKey = pk.toLowerCase().replace(/[\s　]/g, "");
        if (cleanTarget === cleanKey || cleanTarget.indexOf(cleanKey) !== -1 || cleanKey.indexOf(cleanTarget) !== -1) {
          return photoMap[pk];
        }
      }
      return "";
    }

    for (var r = 1; r < rows.length; r++) {
      if (!rows[r][0]) continue;
      var evName = rows[r][2] ? rows[r][2].toString().trim() : "";
      var summaryStr = rows[r][3] ? rows[r][3].toString() : "";
      var photoMap = getEventPhotoMap(evName);

      // 建立商品圖片陣列：[{ name, photo }]
      var productPhotos = [];
      // D欄（現有商品）+ T欄（缺貨商品）都納入圖片對照
      var oosStr = rows[r][19] ? rows[r][19].toString() : "";
      var allProdNames = extractProductNames(summaryStr).concat(extractProductNames(oosStr));
      var seenPhoto = {};
      allProdNames.forEach(function(pn) {
        if (seenPhoto[pn]) return;
        seenPhoto[pn] = true;
        productPhotos.push({ name: pn, photo: findPhoto(photoMap, pn) });
      });

      orders.push({
        id: rows[r][0].toString(),
        date: rows[r][1],
        event: evName,
        summary: summaryStr,
        amount: rows[r][4] || 0,
        shippingFee: rows[r][5] ? Number(rows[r][5]) : 0, // F欄
        payMethod: rows[r][7] ? rows[r][7].toString() : "",
        name: rows[r][8] ? rows[r][8].toString() : "",
        phone: (function() {
          var raw = rows[r][9] ? rows[r][9].toString().trim() : "";
          var pt = rows[r][11] ? rows[r][11].toString() : "";
          var pay = rows[r][7] ? rows[r][7].toString() : "";
          var isTW = pt.indexOf("賣貨便") !== -1 || pay.indexOf("郵局") !== -1;
          // TW 手機號碼應為10位，開頭0被 Sheet 去掉時補回
          if (isTW && raw.length === 9 && raw.charAt(0) !== "0") raw = "0" + raw;
          return raw;
        })(),
        email: rows[r][10] ? rows[r][10].toString() : "",
        pickupType: rows[r][11] ? rows[r][11].toString() : "",
        pickupCode: rows[r][12] ? rows[r][12].toString() : "",
        pickupName: rows[r][13] ? rows[r][13].toString() : "",
        pickupAddress: rows[r][14] ? rows[r][14].toString() : "",
        status: rows[r][15] ? rows[r][15].toString() : "待處理",
        remark: rows[r][16] ? rows[r][16].toString() : "",          // Q欄：備註
        paymentTime: rows[r][17] ? rows[r][17] : "",                // R欄：付款確認時間
        shipmentRef:   rows[r][18] ? rows[r][18].toString() : "",  // S欄：運單號碼(HK)/賣貨便網址(TW)
        stockoutItems: rows[r][19] ? rows[r][19].toString() : "", // T欄：Stockout Items
        stockoutAmt:   rows[r][20] ? Number(rows[r][20]) : 0,       // U欄：Stockout Amount
        stockoutJpy:   rows[r][21] ? Number(rows[r][21]) : 0,       // V欄：Stockout JPY
        weight:        rows[r][22] ? Number(rows[r][22]) : 0,     // W欄：重量(g)
        intlFee:       rows[r][23] ? Number(rows[r][23]) : 0,     // X欄：國際運費(NT$)
        arrivalPhoto:  rows[r][23] ? rows[r][23].toString() : "", // X欄：到貨圖片URL
        cm711OrderId:  rows[r][24] ? rows[r][24].toString() : "", // Y欄：711 Order ID
        recipientName: rows[r][25] ? rows[r][25].toString() : "", // Z欄：Recipient Name
        storeName:     rows[r][26] ? rows[r][26].toString() : "", // AA欄：Store Name
        refundPhoto:   rows[r][27] ? rows[r][27].toString() : "", // AB欄：RefundPhoto
        refundBankName:rows[r][29] ? rows[r][29].toString() : "", // AD欄：銀行簡稱
        refundBankCode:rows[r][30] ? rows[r][30].toString() : "", // AE欄：匯款銀行代碼
        refundBankAcc: rows[r][31] ? rows[r][31].toString() : "", // AF欄：銀行帳號
        refundAccName: rows[r][32] ? rows[r][32].toString() : "", // AG欄：受款人戶名
        refundAccName: rows[r][32] ? rows[r][32].toString() : "", // AG欄：受款人戶名
        productPhotos: productPhotos  // ← 商品圖片對照
      });
    }
    
    if (action === "getAdminOrders") {
      return ContentService.createTextOutput(JSON.stringify({ orders: orders }))
                            .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify(orders))
                            .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 🍔 【分流 3】前台選單：獲取商品分頁清單
  if (action === "getMenuList") {
    var sheets = ss.getSheets();
    var menuList = [];
    for (var i = 0; i < sheets.length; i++) {
      var name = sheets[i].getName();
      if (!name.startsWith("[Data]") && name !== "訂單紀錄" && name !== "易寄取地址" && name !== "Blank" && name !== "收單截止時間") {
        menuList.push(name);
      }
    }
    return ContentService.createTextOutput(JSON.stringify(menuList)).setMimeType(ContentService.MimeType.JSON);
  }

  // 🕐 【分流 5】讀取所有 event 的截止時間
  if (action === "getDeadlines") {
    var result = syncAndGetDeadlines(ss);
    return ContentService.createTextOutput(JSON.stringify(result))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "getNissenProduct") {
    var nissenResult = fetchNissenProduct_(param.url || '');
    return ContentService.createTextOutput(JSON.stringify(nissenResult))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // 代攞商品相。瀏覽器直接畫日本網站啲相落 canvas 會整污糟咗塊 canvas，
  // 之後匯出唔到；由呢邊攞返轉做 data URI 就冇呢個問題。
  if (action === "proxyImage") {
    var imgUrl = param.url || '';
    if (!imgUrl) {
      return jsonpOrJson_(param, { error: '缺少圖片網址' });
    }
    try {
      var imgResp = UrlFetchApp.fetch(imgUrl, {
        headers: {
          'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
          'Accept': 'image/avif,image/webp,image/*,*/*;q=0.8',
          'Referer': imgUrl
        },
        muteHttpExceptions: true,
        followRedirects: true
      });
      if (imgResp.getResponseCode() !== 200) {
        return jsonpOrJson_(param, { error: '攞唔到圖片 (HTTP ' + imgResp.getResponseCode() + ')' });
      }
      var imgBlob = imgResp.getBlob();
      var imgBytes = imgBlob.getBytes();
      // base64 之後大三分一，太大會爆 GAS 個回應上限
      if (imgBytes.length > 4 * 1024 * 1024) {
        return jsonpOrJson_(param, { error: '張相太大（' + Math.round(imgBytes.length / 1024) + 'KB）' });
      }
      var imgType = imgBlob.getContentType() || 'image/jpeg';
      return jsonpOrJson_(param, {
        dataUri: 'data:' + imgType + ';base64,' + Utilities.base64Encode(imgBytes)
      });
    } catch (imgErr) {
      return jsonpOrJson_(param, { error: '攞圖片失敗：' + imgErr.toString() });
    }
  }

  if (action === "getNetseaProduct") {
    var nsResult = fetchNetseaProduct_(param.url || '');
    return ContentService.createTextOutput(JSON.stringify(nsResult))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "getZozoProduct") {
    var zozoResult = fetchZozoProduct_(param.url || '');
    return ContentService.createTextOutput(JSON.stringify(zozoResult))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "getKhw1Product") {
    var khw1Result = fetchKhw1Product_(param.url || '');
    // jsonpOrJson_ 兩用：post.html 直接 fetch() 攞返 plain JSON；
    // admin.html 用 gasJsonp()（帶 callback 參數）就包成 JSONP，
    // 唔使淨係為呢個 action 另開一條路。
    return jsonpOrJson_(param, khw1Result);
  }

  if (action === "getNissenFeatured") {
    // 1. 讀取每日快取（最快，不佔 quota）
    var campItems = getNissenCampaignCached_();
    if (!campItems || campItems.length === 0) {
      // 2. 快取為空（首次），即時抓取並同時寫入快取
      campItems = fetchNissenCampaign_();
      if (campItems && campItems.length > 0) {
        try {
          var props = PropertiesService.getScriptProperties();
          props.setProperty('nissen_campaign_items',   JSON.stringify(campItems));
          props.setProperty('nissen_campaign_updated', new Date().toISOString());
        } catch(e) {}
      }
    }
    if (campItems && campItems.length > 0) {
      return ContentService.createTextOutput(JSON.stringify({ items: campItems }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    // 3. 備用：讀取「Nissen精選」試算表
    var featSheet = ss.getSheetByName("Nissen精選");
    if (!featSheet) {
      return ContentService.createTextOutput(JSON.stringify({ items: [] }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    var featRows = featSheet.getDataRange().getValues();
    var featItems = [];
    for (var fi = 1; fi < featRows.length; fi++) {
      if (!featRows[fi][0] && !featRows[fi][2]) continue;
      featItems.push({
        name:    featRows[fi][0] ? featRows[fi][0].toString() : '',
        image:   featRows[fi][1] ? featRows[fi][1].toString() : '',
        url:     featRows[fi][2] ? featRows[fi][2].toString() : '',
        hkPrice: featRows[fi][3] ? featRows[fi][3].toString() : '',
        twPrice: featRows[fi][4] ? featRows[fi][4].toString() : ''
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ items: featItems }))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // 🛍️ 【分流 6】後台：讀取指定 event 的商品清單（含 stockLimit）
  if (action === "getProducts") {
    var sheetName = param.sheetName ? param.sheetName : "";
    if (!sheetName) {
      return ContentService.createTextOutput(JSON.stringify({ error: "缺少 sheetName 參數" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    var pSheet = ss.getSheetByName(sheetName);
    if (!pSheet) {
      return ContentService.createTextOutput(JSON.stringify({ error: "找不到分頁：" + sheetName }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    var pRows = pSheet.getDataRange().getValues();
    var productList = [];
    for (var pi = 1; pi < pRows.length; pi++) {
      if (!pRows[pi][0]) continue;
      var rawLmt = pRows[pi][9]; // J欄 subStockLimit 文字
      var subsRaw = pRows[pi][1] ? pRows[pi][1].toString() : "";
      productList.push({
        row: pi + 1,
        name: pRows[pi][0] ? pRows[pi][0].toString() : "",
        subs: subsRaw,
        subsArray: subsRaw.split(" / ").map(function(s){ return s.trim(); }),
        yen: parseFloat(pRows[pi][2]) || 0,
        cost: parseFloat(pRows[pi][3]) || 0,
        price: parseFloat(pRows[pi][4]) || 0,
        photo: pRows[pi][5] ? pRows[pi][5].toString() : "",
        weight: parseFloat(pRows[pi][6]) || 0,
        quote: pRows[pi][7] ? pRows[pi][7].toString() : "",
        twd: parseFloat(pRows[pi][8]) || 0,
        subStockLimit: parseSubStockLimit(rawLmt)  // { 款式: 庫存 } 或 null
      });
    }
    return ContentService.createTextOutput(JSON.stringify({ sheetName: sheetName, products: productList }))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // 🔄 【分流 9】同步購貨紀錄 E/F 欄（遷移/重建用）
  if (action === "syncPurchaseRecord") {
    var evName = param.event ? param.event : "";
    if (!evName) {
      return ContentService.createTextOutput(JSON.stringify({ error: "缺少 event 參數" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    initSyncPurchaseRecord(ss, evName);
    return ContentService.createTextOutput(JSON.stringify({ result: "success", event: evName }))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // 📈 【分流 8】盈利頁面資料
  if (action === "getProfitData") {
    var evName = param.event ? param.event : "";
    if (!evName) {
      return ContentService.createTextOutput(JSON.stringify({ error: "缺少 event 參數" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    try {
      var profitResult = getProfitData(ss, evName);
      return ContentService.createTextOutput(JSON.stringify(profitResult))
                           .setMimeType(ContentService.MimeType.JSON);
    } catch(profitErr) {
      Logger.log("getProfitData 錯誤: " + profitErr.toString());
      return ContentService.createTextOutput(JSON.stringify({ error: profitErr.toString() }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 📦 【分流 10】讀取所有 event 的已購買數量（購貨紀錄 G欄）
  if (action === "getPurchasedQty") {
    // 讀取所有 sheets：舊版總表 + 所有 per-event sheets
    var result2 = {};
    var allSheets2 = ss.getSheets();
    for (var si2 = 0; si2 < allSheets2.length; si2++) {
      var sName2 = allSheets2[si2].getName();
      // 舊版總表
      if (sName2 === "購貨紀錄" || sName2 === "[Data]購貨紀錄") {
        var pRows2 = allSheets2[si2].getDataRange().getValues();
        for (var pr2 = 1; pr2 < pRows2.length; pr2++) {
          if (!pRows2[pr2][0]) continue;
          var ev2   = pRows2[pr2][0].toString().trim();
          var nm2   = pRows2[pr2][1] ? pRows2[pr2][1].toString().trim() : "";
          var sb2   = pRows2[pr2][2] ? pRows2[pr2][2].toString().trim() : "";
          var qty2  = Number(pRows2[pr2][6]) || 0;
          if (!nm2) continue;
          if (!result2[ev2]) result2[ev2] = {};
          result2[ev2][nm2 + "|" + sb2] = qty2;
        }
      }
      // per-event sheet
      var pem2 = sName2.match(/^\[Data\]\((.+)\)購貨紀錄$/);
      if (pem2) {
        var ev2b = pem2[1];
        if (!result2[ev2b]) result2[ev2b] = {};
        var peRows2 = allSheets2[si2].getDataRange().getValues();
        for (var per2 = 1; per2 < peRows2.length; per2++) {
          if (!peRows2[per2][0]) continue;
          var pnm2 = peRows2[per2][0].toString().trim();
          var psb2 = peRows2[per2][1] ? peRows2[per2][1].toString().trim() : "";
          var pqt2 = Number(peRows2[per2][5]) || 0; // F(index5)=已購買
          if (!pnm2) continue;
          result2[ev2b][pnm2 + "|" + psb2] = pqt2;
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify(result2))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  // 📊 【分流 7】訂貨區：讀取指定 event 已付款訂單統計
  if (action === "getOrderSummary") {
    var evName = param.event ? param.event : "";
    if (!evName) {
      return ContentService.createTextOutput(JSON.stringify({ error: "缺少 event 參數" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    try {
      var summary = getOrderSummary(ss, evName);
      // getOrderSummary 回傳 { items, batches }
      var summaryItems   = summary.items   || summary || [];
      var summaryBatches = summary.batches || [];
      return ContentService.createTextOutput(JSON.stringify({ event: evName, items: summaryItems, batches: summaryBatches }))
                           .setMimeType(ContentService.MimeType.JSON);
    } catch(sumErr) {
      Logger.log("getOrderSummary 錯誤: " + sumErr.toString());
      return ContentService.createTextOutput(JSON.stringify({ error: sumErr.toString(), items: [], batches: [] }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // ── Nissen 訂單通知（前台下單後 fire-and-forget，通知 admin）──
  if (action === "nissenOrderNotify") {
    try {
      var orderStr = param.orderJson ? param.orderJson : "{}";
      var nOrder = {};
      try { nOrder = JSON.parse(orderStr); } catch(pe) {}
      if (nOrder.order_no || nOrder.orderNo) {
        var nSubject = "[新 Nissen 代購訂單] " + (nOrder.order_no || nOrder.orderNo) + " — " + (nOrder.customer_name || nOrder.customerName || "") + " (" + (nOrder.region || "").toUpperCase() + ")";
        var nBody = buildNissenEmailHtml_(nOrder, false);
        MailApp.sendEmail({ to: MY_NOTIFICATION_EMAIL, subject: nSubject, htmlBody: nBody });
      }
    } catch(ne) { Logger.log("nissenOrderNotify error: " + ne); }
    return jsonpOrJson_(param, { result: "ok" });
  }

  if (action === "getOrderSummaryAll") {
    try {
      return ContentService.createTextOutput(JSON.stringify({ events: getOrderSummaryAll(ss) }))
                           .setMimeType(ContentService.MimeType.JSON);
    } catch (allErr) {
      Logger.log("getOrderSummaryAll 錯誤: " + allErr);
      return ContentService.createTextOutput(JSON.stringify({ error: allErr.toString(), events: [] }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 認唔到嘅指令要即刻講明，唔好靜靜跌落分流 4。
  // 跌落去嘅話會慢慢掃勻成個訂單紀錄，再回傳一份商品清單 ——
  // 呼叫方等到 timeout 都等唔到佢要嘅欄位，睇落就好似「一直查詢中」。
  // 呢個情況通常代表 GAS 未重新部署，即係部署緊嘅版本仲未有呢個指令。
  if (action && action !== "getItems") {
    return jsonpOrJson_(param, {
      result: "error",
      error: "unknown_action",
      action: action,
      message: "呢個 GAS 版本未認得指令「" + action + "」。請喺 Apps Script 撳「部署 → 管理部署 → ✏️ → 版本：新版本 → 部署」，千祈唔好用「新增部署」（會換咗網址）。"
    });
  }

  // 🛒 【分流 4】下單前台：兼容 action=getItems、event=、sheetName= 三種方式
  var targetSheetName = param.event ? param.event : (param.sheetName ? param.sheetName : "");
  
  if (!targetSheetName) {
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var n = sheets[i].getName();
      if (!n.startsWith("[Data]") && n !== "訂單紀錄" && n !== "易寄取地址" && n !== "Blank" && n !== "收單截止時間") {
        targetSheetName = n;
        break;
      }
    }
  }

  var productSheet = ss.getSheetByName(targetSheetName);
  var products = [];
  if (productSheet) {
    var productRows = productSheet.getDataRange().getValues();
    for (var i = 1; i < productRows.length; i++) {
      if (!productRows[i][0]) continue;
      var rawLimit = productRows[i][9]; // J欄：stockLimit
      products.push({
        id: "p_" + i,
        name: productRows[i][0],
        subs: String(productRows[i][1]).split(" / ").map(function(s) { return s.trim(); }),
        yen: productRows[i][2],
        cost: productRows[i][3],
        price: parseFloat(productRows[i][4]) || 0,
        photo: productRows[i][5],
        weight: productRows[i][6],
        quote: productRows[i][7],
        twd: parseFloat(productRows[i][8]) || 0,        // I欄：台幣價格
        subStockLimit: parseSubStockLimit(rawLimit)  // J欄：{ 款式: 庫存 } 或 null
      });
    }
  }
  
  var addressSheet = ss.getSheetByName("易寄取地址");
  var addresses = [];
  if (addressSheet) {
    var addressRows = addressSheet.getDataRange().getValues();
    for (var j = 1; j < addressRows.length; j++) {
      if (!addressRows[j][1]) continue;
      addresses.push({
        type: addressRows[j][0], code: addressRows[j][1].toString(), name: addressRows[j][2],
        region: addressRows[j][3], district: addressRows[j][4], address: addressRows[j][5]
      });
    }
  }
  
  var deadlineData = syncAndGetDeadlines(ss);
  var thisDeadline = deadlineData[targetSheetName] || "";

  // 📦 統計這個 event 的已購數量（HK+TW 合計，每個「商品名|款式」各自計算）
  var orderedQty = getOrderedQtyForEvent(ss, targetSheetName);

  return ContentService.createTextOutput(JSON.stringify({
    currentSheet: targetSheetName,
    products: products,
    addresses: addresses,
    deadline: thisDeadline,
    orderedQty: orderedQty   // { "商品名|款式": 已購數量 }
  })).setMimeType(ContentService.MimeType.JSON);
}

// =================================================================
// Nissen 商品資料抓取
// =================================================================
function fetchNissenProduct_(url) {
  if (!url || url.indexOf('nissen.co.jp') === -1) {
    return { error: '請輸入有效的 Nissen 商品網址 (nissen.co.jp)' };
  }
  try {
    var opts = {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ja,en;q=0.5'
      },
      muteHttpExceptions: true,
      followRedirects: true
    };
    var resp = UrlFetchApp.fetch(url, opts);
    var code = resp.getResponseCode();
    if (code !== 200) {
      return { error: '無法讀取商品頁面 (HTTP ' + code + ')。請確認網址是否正確。' };
    }
    var html = resp.getContentText('UTF-8');
    var result = { name: '', image: null, price: null, priceMax: null, variants: [], rawUrl: url };

    // 1. 商品名稱: var itemName = "..."
    var nameM = html.match(/var\s+itemName\s*=\s*"([^"]+)"/);
    if (nameM) {
      result.name = decodeJsString_(nameM[1]);
    } else {
      var ogM = html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)["']/i)
             || html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:title["']/i);
      if (ogM) result.name = ogM[1].trim();
    }

    // 1b. 商品圖片: og:image
    var imgM = html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i)
            || html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:image["']/i);
    if (imgM) result.image = imgM[1];

    // 2. 税込価格: var priceL = 最低価格, var price = 最高価格
    var pLM = html.match(/var\s+priceL\s*=\s*"(\d+)"/);
    var pMaxM = html.match(/var\s+price\s*=\s*"(\d+)"/);
    if (pLM) result.price = parseInt(pLM[1]);
    if (pMaxM) {
      var pMax = parseInt(pMaxM[1]);
      if (result.price && pMax !== result.price) result.priceMax = pMax;
      else if (!result.price) result.price = pMax;
    }

    // 3. SKU一覧: var skuList = [{...}, ...]
    var skuJson = nissenExtractArray_(html, 'skuList');
    if (skuJson) {
      try {
        var skuList = JSON.parse(skuJson);
        var colorMap = {};
        var colorOrder = [];
        skuList.forEach(function(sku) {
          var cName = (sku.exhibitionColorName || sku.colorName || '').trim();
          var sName = (sku.exhibitionSizeName  || sku.sizeName  || '').trim();
          if (!cName && !sName) return;
          if (!colorMap[cName]) { colorMap[cName] = []; colorOrder.push(cName); }
          var inStock = !sku.soldOut && (typeof sku.allocatableStock === 'undefined' || sku.allocatableStock > 0);
          colorMap[cName].push({ name: sName, inStock: inStock, reserve: !!sku.reserve, price: sku.retailPrice ? parseInt(sku.retailPrice) : null });
        });
        colorOrder.forEach(function(c) {
          result.variants.push({ group: c || '款式', options: colorMap[c] });
        });
      } catch(e3) {}
    }

    // 翻譯日文→繁中（一次過譯晒，唔好逐個字串叫一次）
    translateProductInPlace_(result);

    return result;
  } catch(e) {
    return { error: '抓取失敗：' + e.toString() };
  }
}

// =================================================================
// ZOZOTOWN 商品抓取
// ZOZOTOWN 係 Next.js，成個商品資料都喺 <script id="__NEXT_DATA__"> 入面，
// 唔使拆 HTML，直接讀 JSON 就攞到名／稅込價／圖／顏色尺碼／庫存。
// =================================================================
function fetchZozoProduct_(url) {
  if (!url || url.indexOf('zozo.jp') === -1) {
    return { error: '請輸入有效的 ZOZOTOWN 商品網址 (zozo.jp)' };
  }
  try {
    var resp = UrlFetchApp.fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ja,en;q=0.5'
      },
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (resp.getResponseCode() !== 200) {
      return { error: '無法讀取商品頁面 (HTTP ' + resp.getResponseCode() + ')。請確認網址是否正確。' };
    }
    var html = resp.getContentText('UTF-8');
    var result = { name: '', image: null, price: null, priceMax: null, variants: [], rawUrl: url };

    // ── 主要路徑：__NEXT_DATA__ ──
    var m = html.match(/<script id="__NEXT_DATA__"[^>]*>([\s\S]*?)<\/script>/);
    if (m) {
      try {
        var d = JSON.parse(m[1]);
        var fsr = d && d.props && d.props.pageProps && d.props.pageProps.frontServerResult;
        if (fsr && fsr.goods) {
          var g = fsr.goods;
          result.name  = (g.goodsName || '').trim();
          result.image = g.defaultImageUrl || null;

          var pi = g.priceInfo || {};
          if (pi.price) result.price = parseInt(pi.price);          // 稅込售價
          if (pi.doublePriceLabel && pi.doublePriceLabel.price) {
            result.origPrice = parseInt(pi.doublePriceLabel.price); // 原價
          }
          if (pi.discountRate) result.discountRate = pi.discountRate;

          // 品牌／店舖，方便你落單時知去邊間鋪
          if (fsr.brand) result.brand = fsr.brand.brandName || '';
          if (fsr.shop)  result.shop  = fsr.shop.shopNameEn || '';

          // 顏色 → 尺碼，並標明有冇貨
          var shelves = (fsr.goodsShelfInfo && fsr.goodsShelfInfo.shelves) || [];
          var colorMap = {}, colorOrder = [];
          shelves.forEach(function(s) {
            var c = (s.colorName || '').trim();
            var z = (s.sizeShortName || s.sizeName || '').trim();
            if (!c && !z) return;
            if (!colorMap[c]) { colorMap[c] = []; colorOrder.push(c); }
            colorMap[c].push({
              name: z,
              inStock: s.captionType === 'INSTOCK',
              reserve: false,
              price: null
            });
          });
          colorOrder.forEach(function(c) {
            result.variants.push({ group: c || '款式', options: colorMap[c] });
          });
        }
      } catch(pe) { /* 解析唔到就落去 fallback */ }
    }

    // ── 後備路徑：og / JSON-LD（萬一 ZOZO 改咗結構或者擋咗）──
    if (!result.name) {
      var og = html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)["']/i);
      if (og) result.name = og[1].replace(/｜.*$/, '').replace(/^【セール】/, '').trim();
    }
    if (!result.image) {
      var oi = html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i);
      if (oi) result.image = oi[1];
    }
    if (!result.price) {
      var ld = html.match(/"price"\s*:\s*"?(\d+)"?/);
      if (ld) result.price = parseInt(ld[1]);
    }
    if (!result.name && !result.price) {
      return { error: 'ZOZOTOWN 讀唔到商品資料，可能係網址唔啱或者網站暫時擋咗。請稍後再試。' };
    }

    // 翻譯日文→繁中：一次過譯晒，唔好逐個字串叫一次
    // （一件貨有七、八隻顏色就要等七、八次翻譯，前台會一直「查詢中」）
    translateProductInPlace_(result);

    return result;
  } catch(e) {
    return { error: '抓取失敗：' + e.toString() };
  }
}

// 收集商品名同所有款式名，合併成一次 LanguageApp 呼叫再派返落去
// 公仔名字典。出 post 唔會用日文原名，而 LanguageApp 譯人物名多數譯到唔知咩嚟，
// 所以直接查表。長嘅要行先，唔係「マイメロ」會搶咗「マイメロディ」。
var CHARACTERS_ = [
  // [日文原名, 台灣叫法, 香港叫法]
  // 香港多數直接叫英文名，台灣用中文譯名，所以兩邊分開存。
  // 長嘅要行先，唔係「マイメロ」會搶咗「マイメロディ」。
  ['マイメロディ', '美樂蒂', 'My Melody'], ['マイメロ', '美樂蒂', 'My Melody'],
  ['クロミ', '酷洛米', 'Kuromi'],
  ['ハローキティ', 'Hello Kitty', 'Hello Kitty'], ['キティ', 'Hello Kitty', 'Hello Kitty'],
  ['シナモロール', '大耳狗', 'Cinnamoroll'], ['シナモン', '大耳狗', 'Cinnamoroll'],
  ['ポムポムプリン', '布丁狗', 'Pompompurin'],
  ['ぐでたま', '蛋黃哥', 'Gudetama'],
  ['けろけろけろっぴ', '大眼蛙', 'Keroppi'], ['けろっぴ', '大眼蛙', 'Keroppi'],
  ['リトルツインスターズ', '雙子星', 'Little Twin Stars'],
  ['キキララ', '雙子星', 'Little Twin Stars'], ['キキ&ララ', '雙子星', 'Little Twin Stars'],
  ['ハンギョドン', '人魚漢頓', 'Hangyodon'],
  ['バッドばつ丸', '酷企鵝', 'Badtz-Maru'], ['ばつ丸', '酷企鵝', 'Badtz-Maru'],
  ['ポチャッコ', '帕恰狗', 'Pochacco'],
  ['タキシードサム', '山姆企鵝', 'Tuxedosam'],
  ['あひるのペックル', '貝克鴨', 'Pekkle'],
  ['こぎみゅん', '小麥狗', 'Kogimyun'],
  ['ウィッシュミーメル', '許願兔', 'Wish me mell'],
  ['ちいかわ', '吉伊卡哇', 'Chiikawa'],
  ['ハチワレ', '小八', 'Hachiware'],
  ['モモンガ', '鼯鼠', 'Momonga'],
  ['くりまんじゅう', '栗子饅頭', 'Kurimanju'],
  ['すみっコぐらし', '角落生物', 'Sumikko Gurashi'],
  ['コリラックマ', '小白熊', 'Korilakkuma'],
  ['キイロイトリ', '黃小鳥', 'Kiiroitori'],
  ['リラックマ', '拉拉熊', 'Rilakkuma'],
  ['ミッフィー', '米飛兔', 'Miffy'],
  ['スヌーピー', '史努比', 'Snoopy'],
  ['ムーミン', '嚕嚕米', 'Moomin'],
  ['ドラえもん', '哆啦A夢', '多啦A夢'],
  ['ピカチュウ', '皮卡丘', '比卡超'],
  ['ポケモン', '寶可夢', 'Pokémon'],
  ['パペットスンスン', '啾啾', 'SunSun'],
  ['くまのプーさん', '小熊維尼', 'Winnie the Pooh'],
  ['プーさん', '小熊維尼', 'Winnie the Pooh'],
  ['スティッチ', '史迪奇', 'Stitch'],
  ['トトロ', '龍貓', '龍貓'],
  ['サンリオ', '三麗鷗', 'Sanrio']
];

// 回台灣同香港兩個叫法，邊個地區出 post 就由前端揀
function detectCharacter_(text) {
  if (!text) return null;
  for (var i = 0; i < CHARACTERS_.length; i++) {
    if (text.indexOf(CHARACTERS_[i][0]) !== -1) {
      return { tw: CHARACTERS_[i][1], hk: CHARACTERS_[i][2] };
    }
  }
  return null;
}

// NETSEA 商品頁。出 post 只需要商品名、相、同埋「メーカー希望小売価格」——
// 卸價唔會出街，所以呢度唔攞。
function fetchNetseaProduct_(url) {
  if (!url || url.indexOf('netsea.jp') === -1) {
    return { error: '請輸入有效的商品網址' };
  }
  try {
    var resp = UrlFetchApp.fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ja,en;q=0.5'
      },
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (resp.getResponseCode() !== 200) {
      return { error: '無法讀取商品頁面 (HTTP ' + resp.getResponseCode() + ')' };
    }
    var html = resp.getContentText('UTF-8');
    var result = { name: '', image: null, price: null, variants: [], rawUrl: url };

    var ogT = html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)["']/i);
    if (ogT) result.name = decodeJsString_(ogT[1]);
    var ogI = html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i);
    if (ogI) result.image = ogI[1];

    // 上代（メーカー希望小売価格）寫成「3,500円/点（税抜）」，
    // 卸價寫成「2,415円（税抜）」—— 靠「円/点」分得出邊個係邊個。
    var rp = html.match(/([0-9][0-9,]*)\s*<span[^>]*class=["']taxUnit["'][^>]*>\s*円\s*\/\s*点/);
    if (rp) result.retailYen = parseInt(rp[1].replace(/,/g, ''), 10);

    // 公仔名由日文原名度認 —— 譯完就搵唔返
    var ch = detectCharacter_(result.name);
    result.character   = ch ? ch.tw : '';
    result.characterHk = ch ? ch.hk : '';

    if (!result.name) return { error: '攞唔到商品資料，請確認網址' };
    if (!result.retailYen) {
      result.warn = '搵唔到「メーカー希望小売価格」，價錢要自己填';
    }

    translateProductInPlace_(result);
    return result;
  } catch (e) {
    return { error: '抓取失敗：' + e.toString() };
  }
}

// =================================================================
// khw1.com（世界一百貨批發）商品抓取
// khw1.com 用 Shopline 平台，商品頁內嵌成個商品物件喺
// app.value('product', JSON.parse('...')) 入面，比 JSON-LD 更齊全
// （連款式/庫存/預購說明都有），照住抓取階層：內嵌 JSON → JSON-LD →
// og meta 逐層 fallback。khw1 賣嘅係台灣批發現貨，內容本身已經係
// 繁體中文，唔好叫 translateProductInPlace_（佢個「係咪日文」判斷式
// 連中文都會誤判，白白燒 LanguageApp 配額之餘仲有機會譯錯字）。
// =================================================================
function fetchKhw1Product_(url) {
  if (!url || url.indexOf('khw1.com') === -1) {
    return { error: '請輸入有效的商品網址 (khw1.com)' };
  }
  try {
    var resp = UrlFetchApp.fetch(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'zh-TW,zh;q=0.9'
      },
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (resp.getResponseCode() !== 200) {
      return { error: '無法讀取商品頁面 (HTTP ' + resp.getResponseCode() + ')' };
    }
    var html = resp.getContentText('UTF-8');
    var result = { name: '', image: null, price: null, variants: [], rawUrl: url };

    // ── 主要路徑：Shopline 頁面內嵌的 product 物件 ──
    var product = extractEmbeddedJsonParse_(html, /app\.value\('product',\s*JSON\.parse\('/);
    if (product) {
      result.name = (product.title_translations &&
        (product.title_translations['zh-hant'] || product.title_translations['zh-tw'])) || '';

      result.image = (product.media && product.media[0] && product.media[0].images && product.media[0].images.original)
        ? product.media[0].images.original.url
        : (product.cover_media_array && product.cover_media_array[0] ? product.cover_media_array[0].original_image_url : null);

      // price_sale.cents > 0 先代表有做緊優惠價，否則用原價
      var basePrice = (product.price_sale && product.price_sale.cents > 0) ? product.price_sale : product.price;
      if (basePrice) result.price = basePrice.dollars;

      if (product.field_titles && product.field_titles.length && product.variations && product.variations.length) {
        var groupName = (product.field_titles[0].name_translations &&
          product.field_titles[0].name_translations['zh-hant']) || '款式';
        var options = product.variations.map(function(v) {
          var optName = (v.fields_translations && v.fields_translations['zh-hant'] && v.fields_translations['zh-hant'][0]) || '';
          // 預購/現貨都算「有得買」：quantity 淨係現貨庫存，khw1 好多商品係預購(quantity=0)但一樣落到單
          var inStock = (v.quantity > 0) || product.out_of_stock_orderable === true || product.unlimited_quantity === true;
          return { name: optName, inStock: inStock };
        });
        result.variants.push({ group: groupName, options: options });
      }

      // 預購到貨日說明（例如「預計10月初-10月中到貨」），冇就唔加呢個欄位
      if (product.is_preorder && product.preorder_note_translations && product.preorder_note_translations['zh-hant']) {
        result.note = product.preorder_note_translations['zh-hant'];
      }
    }

    // ── 備援 1：JSON-LD Product schema ──
    if (!result.name) {
      var ldBlocks = html.match(/<script type="application\/ld\+json">[\s\S]*?<\/script>/g) || [];
      for (var li = 0; li < ldBlocks.length; li++) {
        var inner = ldBlocks[li].replace(/^<script type="application\/ld\+json">/, '').replace(/<\/script>$/, '');
        try {
          var ld = JSON.parse(inner);
          if (ld && ld['@type'] === 'Product') {
            result.name = ld.name || '';
            if (ld.image && ld.image.length) result.image = ld.image[0];
            if (ld.offers && ld.offers.price) result.price = ld.offers.price;
            break;
          }
        } catch (eLd) { /* 呢個 script block 唔係有效 JSON，跳過 */ }
      }
    }

    // ── 備援 2：Open Graph meta ──
    if (!result.name) {
      var ogT = html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)["']/i);
      if (ogT) result.name = decodeJsString_(ogT[1]);
    }
    if (!result.image) {
      var ogI = html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i);
      if (ogI) result.image = ogI[1];
    }

    if (!result.name) return { error: '攞唔到商品資料，請確認網址' };

    return result;
  } catch (e) {
    return { error: '抓取失敗：' + e.toString() };
  }
}

function translateProductInPlace_(result) {
  var slots = [];   // { get, set }
  if (result.name) slots.push({ v: result.name, set: function(t){ result.name = t; } });
  (result.variants || []).forEach(function(g) {
    slots.push({ v: g.group, set: function(t){ g.group = t; } });
    (g.options || []).forEach(function(o) {
      if (o.name) slots.push({ v: o.name, set: function(t){ o.name = t; } });
    });
  });

  var JP = /[぀-ヿ一-鿿＀-￯]/;
  var need = [];
  slots.forEach(function(s, i) { if (s.v && JP.test(s.v)) need.push(i); });
  if (!need.length) return;

  try {
    var joined = need.map(function(i){ return slots[i].v; }).join('\n');
    var out = LanguageApp.translate(joined, 'ja', 'zh-TW') || '';
    var parts = out.split('\n');
    if (parts.length === need.length) {
      need.forEach(function(idx, k) {
        var t = (parts[k] || '').trim();
        if (t) slots[idx].set(t);
      });
      return;
    }
  } catch(e) { /* 跌落去逐個譯 */ }

  // 一次過譯有時會併行或者拆行，行數對唔返。
  // 以前呢個情況會成件事保留日文 —— 一件貨有幾十個尺碼顏色，
  // 差一行就連商品名都唔譯。所以對唔返就逐個譯，慢啲但唔會全軍覆沒。
  need.forEach(function(idx) {
    try {
      var t = LanguageApp.translate(slots[idx].v, 'ja', 'zh-TW');
      if (t && t.trim()) slots[idx].set(t.trim());
    } catch(e) { /* 呢個譯唔到就保留日文 */ }
  });
}

// Nissen 個商品名係由頁面嘅 JS 字面值 var itemName = "…" 抽出嚟，
// 入面啲 （ \/ \" 係未解過碼嘅轉義，照抄出去就會見到成串 （。
function decodeJsString_(s) {
  if (!s) return s;
  return String(s)
    .replace(/\\u([0-9a-fA-F]{4})/g, function(_, h) {
      return String.fromCharCode(parseInt(h, 16));
    })
    .replace(/\\x([0-9a-fA-F]{2})/g, function(_, h) {
      return String.fromCharCode(parseInt(h, 16));
    })
    .replace(/\\n/g, ' ')
    .replace(/\\t/g, ' ')
    .replace(/\\(["'\/\\])/g, '$1')
    .trim();
}

// 由 HTML 攞出 `<prefixRegex>'<JS 轉義過嘅 JSON>'` 呢種 pattern 嵌入嘅資料
// （例如 Shopline 頁面嘅 app.value('product', JSON.parse('...'))）。
// JS 單引號字串入面嘅轉義規則（\uXXXX 深度、邊個字要轉義）好易同手寫
// regex 想像唔一樣，最穩陣係逐個字元行過條轉義路徑抽出原始（仍轉義住嘅）
// 字串內容，再交返俾 JS 引擎自己 eval 做一次「當呢個係字串常量」嘅解碼
// （同瀏覽器行為一致），先至 JSON.parse 出物件。
function extractEmbeddedJsonParse_(html, prefixRegex) {
  var m = html.match(prefixRegex);
  if (!m) return null;
  var i = m.index + m[0].length; // 指住緊接住開頭單引號之後嗰個字元
  var buf = [];
  while (i < html.length) {
    var ch = html.charAt(i);
    if (ch === '\\') {
      buf.push(ch, html.charAt(i + 1));
      i += 2;
      continue;
    }
    if (ch === "'") break;
    buf.push(ch);
    i++;
  }
  try {
    var jsString = eval("'" + buf.join('') + "'");
    return JSON.parse(jsString);
  } catch (e) {
    return null;
  }
}

function nissenTranslate_(text) {
  if (!text || !/[぀-ヿ一-鿿＀-￯]/.test(text)) return text;
  try { return LanguageApp.translate(text, 'ja', 'zh-TW') || text; }
  catch(e) { return text; }
}

// 从 HTML 中提取 var <varName> = [...] 的完整 JSON 数组（支持嵌套括号）
function nissenExtractArray_(html, varName) {
  var re = new RegExp('var\\s+' + varName + '\\s*=\\s*\\[');
  var m = re.exec(html);
  if (!m) return null;
  var i = m.index + m[0].length - 1;
  var depth = 0, inStr = false, strChar = '', esc = false;
  for (; i < html.length; i++) {
    var c = html[i];
    if (esc)                  { esc = false; continue; }
    if (c === '\\' && inStr)  { esc = true;  continue; }
    if (inStr)                { if (c === strChar) inStr = false; continue; }
    if (c === '"' || c === "'") { inStr = true; strChar = c; continue; }
    if (c === '[') { depth++; continue; }
    if (c === ']') { if (--depth === 0) return html.substring(m.index + m[0].length - 1, i + 1); }
  }
  return null;
}

// 抓取 Nissen 官網首頁 campaign 橫幅
function fetchNissenCampaign_() {
  try {
    var resp = UrlFetchApp.fetch('https://www.nissen.co.jp/', {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'ja,en;q=0.5'
      },
      muteHttpExceptions: true,
      followRedirects: true
    });
    if (resp.getResponseCode() !== 200) return [];
    var html = resp.getContentText('UTF-8');

    var items = [], seen = {};

    function normalizeUrl_(u) {
      if (!u) return 'https://www.nissen.co.jp/';
      u = u.trim();
      if (u.indexOf('http') === 0) return u;
      if (u.indexOf('//') === 0)   return 'https:' + u;
      if (u.charAt(0) === '/')     return 'https://www.nissen.co.jp' + u;
      return u;
    }

    // 策略: 搜尋 <a href="...">...<img src 或 data-src="...jpg/png/webp">...</a>
    var reLink = /<a(?:\s[^>]*)?\shref=["']([^"']+)["'][^>]*>([\s\S]{0,3000}?)<\/a>/gi;
    var mL;
    while ((mL = reLink.exec(html)) !== null) {
      if (items.length >= 20) break;
      var href = mL[1].trim();
      if (href.indexOf('javascript') === 0 || href === '#' || href === '') continue;

      var inner = mL[2];
      // 尋找帶有圖片副檔名的 src 或 data-src
      var reImg = /\s(?:data-src|src)=["']([^"']+\.(?:jpg|jpeg|png|webp)(?:\?[^"']*)?)["']/i;
      var mI = reImg.exec(inner);
      if (!mI) continue;
      var imgSrc = mI[1].trim();
      if (!imgSrc || imgSrc.length < 8) continue;

      // 過濾掉明顯的 icon/logo/common 小圖
      if (/\/icon|\/logo|\/common\/|\/btn|\/arrow|\/parts\/|sprite|blank\.png|spacer/i.test(imgSrc)) continue;

      imgSrc = normalizeUrl_(imgSrc);
      if (seen[imgSrc]) continue;

      // 取得 alt 文字
      var mAlt = /\salt=["']([^"']*)["']/i.exec(inner);
      var altText = mAlt ? mAlt[1].trim() : '';

      // 翻譯日文 alt 文字為繁體中文
      var nameZh = altText ? nissenTranslate_(altText.substring(0, 60)) : '';
      seen[imgSrc] = true;
      items.push({
        image:   imgSrc.substring(0, 250),
        url:     normalizeUrl_(href).substring(0, 250),
        name:    nameZh,
        hkPrice: '',
        twPrice: ''
      });
    }

    return items.slice(0, 20);
  } catch(e) {
    return [];
  }
}

// 讀取 PropertiesService 快取
function getNissenCampaignCached_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var cached = props.getProperty('nissen_campaign_items');
    if (cached) return JSON.parse(cached);
  } catch(e) {}
  return null;
}

// 更新快取（每日 trigger 呼叫，或手動執行）
function refreshNissenCampaignCache() {
  var items = fetchNissenCampaign_();
  if (!items || items.length === 0) {
    Logger.log('Nissen campaign: fetch returned empty, cache not updated.');
    return;
  }
  var now = new Date().toISOString();
  var props = PropertiesService.getScriptProperties();

  // 1. 存入 PropertiesService（GAS fallback 用）
  props.setProperty('nissen_campaign_items',   JSON.stringify(items));
  props.setProperty('nissen_campaign_updated', now);

  // 2. 存入 Supabase（前端直接讀，跳過 GAS 冷啟動）
  var supaUrl        = 'https://pksqfpirggvsftvqrtji.supabase.co';
  var supaServiceKey = props.getProperty('SUPA_SERVICE_KEY');
  if (supaServiceKey) {
    try {
      var res = UrlFetchApp.fetch(supaUrl + '/rest/v1/nissen_campaign?id=eq.1', {
        method: 'PATCH',
        headers: {
          'Content-Type':  'application/json',
          'apikey':        supaServiceKey,
          'Authorization': 'Bearer ' + supaServiceKey,
          'Prefer':        'return=minimal'
        },
        payload:           JSON.stringify({ items: items, updated_at: now }),
        muteHttpExceptions: true
      });
      Logger.log('Supabase updated. HTTP ' + res.getResponseCode());
    } catch(e) {
      Logger.log('Supabase update failed: ' + e.message);
    }
  } else {
    Logger.log('SUPA_SERVICE_KEY not set in PropertiesService — skipped Supabase write.');
  }

  Logger.log('Nissen campaign cache updated: ' + items.length + ' items');
}

// 設定每日定時更新（只需運行一次）
function setupNissenCampaignTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'refreshNissenCampaignCache') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('refreshNissenCampaignCache')
    .timeBased()
    .everyDays(1)
    .atHour(3) // 每日凌晨 3 點（時區依 GAS 項目設定，建議設 Asia/Tokyo）
    .create();
  Logger.log('Daily trigger created for refreshNissenCampaignCache.');
}

// =================================================================
// doPost：處理所有前端 POST 請求
// =================================================================
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    var rowData = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ── 功能 1：更新已收款狀態 ──────────────────────────────────────
    if (rowData.action === "updatePaymentStatus") {
      var sheet = ss.getSheetByName("訂單紀錄");
      if (!sheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到『訂單紀錄』分頁" })).setMimeType(ContentService.MimeType.JSON);
      }
      var rows = sheet.getDataRange().getValues();
      var targetRowIndex = -1;
      for (var r = 1; r < rows.length; r++) {
        if (rows[r][0].toString() === rowData.orderId.toString()) {
          targetRowIndex = r + 1;
          break;
        }
      }
      if (targetRowIndex !== -1) {
        sheet.getRange(targetRowIndex, 16).setValue("已收款");
        sheet.getRange(targetRowIndex, 18).setValue(new Date()); // R欄：付款確認時間
        var updatedRows = sheet.getDataRange().getValues();
        var paidRow = updatedRows[targetRowIndex - 1];
        // 自動更新購貨紀錄：F-（未付款）→ E+（已付款）
        try {
          var pItems = parseSummaryItems(paidRow[3] ? paidRow[3].toString() : "");
          var pEvent = paidRow[2] ? paidRow[2].toString().trim() : "";
          updatePurchaseRecordQty(ss, pEvent, pItems, 6, -1); // F欄-
          updatePurchaseRecordQty(ss, pEvent, pItems, 5, 1);  // E欄+
          // ── 即時更新 Supabase qty_paid ──
          updateQtyPaidInSupabase(stripEndPrefix(pEvent), paidRow[3] ? paidRow[3].toString() : "", 1);
        } catch(pe) { Logger.log("確認付款更新購貨紀錄失敗: " + pe.toString()); }
        var customerEmail = sheet.getRange(targetRowIndex, 11).getValue().toString().trim();
        if (customerEmail !== "" && customerEmail.indexOf("@") !== -1) {
          sendPaidEmailToCustomer(ss, paidRow);
        }
        return ContentService.createTextOutput(JSON.stringify({ result: "success" })).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到此訂單編號" })).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    // ── 功能 2：複製 Blank 分頁建立新活動 ────────────────────────────
    if (rowData.action === "createNewEventSheet") {
      var newName = rowData.newSheetName.toString().trim();
      if (ss.getSheetByName(newName)) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "分頁名稱已存在！" })).setMimeType(ContentService.MimeType.JSON);
      }
      var blankSheet = ss.getSheetByName("Blank");
      if (!blankSheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到 Blank 模板分頁" })).setMimeType(ContentService.MimeType.JSON);
      }
      var newSheet = blankSheet.copyTo(ss);
      newSheet.setName(newName);
      return ContentService.createTextOutput(JSON.stringify({ result: "success", message: "成功建立新分頁：" + newName })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ── 功能 3：上架新商品 ────────────────────────────────────────────
    if (rowData.action === "addNewProduct") {
      var tSheet = ss.getSheetByName(rowData.targetSheet) || ss.getSheetByName(stripEndPrefix(rowData.targetSheet));
      if (!tSheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到指定的商品分頁" })).setMimeType(ContentService.MimeType.JSON);
      }
      // A=名稱 B=款式 C=日幣 D=成本HKD E=售價HKD F=圖片 G=重量g H=備註 I=台幣 J=限購
      tSheet.appendRow([
        rowData.p_name || "",
        rowData.p_sub  || "",
        parseFloat(rowData.p_yen)   || 0,
        parseFloat(rowData.p_cost)  || 0,
        parseFloat(rowData.p_price) || 0,
        rowData.p_photo || "",
        parseFloat(rowData.p_weight) || 0,
        rowData.p_quote || "",
        parseFloat(rowData.p_twd)   || 0,
        rowData.p_stockLimit || ""
      ]);
      return ContentService.createTextOutput(JSON.stringify({ result: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 4b：更新現有商品 ─────────────────────────────────────────
    if (rowData.action === "updateProduct") {
      var uSheet = ss.getSheetByName(rowData.targetSheet) || ss.getSheetByName(stripEndPrefix(rowData.targetSheet));
      if (!uSheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到指定的商品分頁" })).setMimeType(ContentService.MimeType.JSON);
      }
      var rowNum = parseInt(rowData.rowNum);
      if (!rowNum || rowNum < 2) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "無效的 rowNum" })).setMimeType(ContentService.MimeType.JSON);
      }
      // A=名稱 B=款式 C=日幣 D=成本HKD E=售價HKD F=圖片 G=重量g H=備註 I=台幣 J=限購
      uSheet.getRange(rowNum, 1, 1, 10).setValues([[
        rowData.p_name  || "",
        rowData.p_sub   || "",
        parseFloat(rowData.p_yen)   || 0,
        parseFloat(rowData.p_cost)  || 0,
        parseFloat(rowData.p_price) || 0,
        rowData.p_photo || "",
        parseFloat(rowData.p_weight) || 0,
        rowData.p_quote || "",
        parseFloat(rowData.p_twd)   || 0,
        rowData.p_stockLimit || ""
      ]]);
      return ContentService.createTextOutput(JSON.stringify({ result: "success" })).setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 5：設定截止時間 ──────────────────────────────────────────
    if (rowData.action === "setDeadline") {
      var dlSheet = ss.getSheetByName("收單截止時間");
      if (!dlSheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到收單截止時間分頁" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var dlRows = dlSheet.getDataRange().getValues();
      for (var r = 1; r < dlRows.length; r++) {
        if (dlRows[r][0].toString() === rowData.eventName) {
          dlSheet.getRange(r + 1, 2).setValue(rowData.deadline);
          return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
                               .setMimeType(ContentService.MimeType.JSON);
        }
      }
      dlSheet.appendRow([rowData.eventName, rowData.deadline]);
      return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 20：儲存退款憑證圖片 URL → AB欄(col28) ──────────────────────
    if (rowData.action === "saveRefundPhoto") {
      var rfSheet = ss.getSheetByName("訂單紀錄");
      if (!rfSheet) return ContentService.createTextOutput(JSON.stringify({result:"error"})).setMimeType(ContentService.MimeType.JSON);
      var rfRows = rfSheet.getDataRange().getValues();
      for (var rfi = 1; rfi < rfRows.length; rfi++) {
        if (rfRows[rfi][0].toString() === (rowData.orderId||"").toString()) {
          rfSheet.getRange(rfi + 1, 28).setValue(rowData.photoUrl || ""); // AB欄
          return ContentService.createTextOutput(JSON.stringify({result:"success"})).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({result:"error",message:"找不到訂單"})).setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 21：發送已退款通知 email ─────────────────────────────────────
    if (rowData.action === "sendRefundNotification") {
      var rnOrderId = rowData.orderId || "";
      var rnSheet   = ss.getSheetByName("訂單紀錄");
      var rnTarget  = null;
      if (rnSheet) {
        var rnRows = rnSheet.getDataRange().getValues();
        for (var rni = 1; rni < rnRows.length; rni++) {
          if (rnRows[rni][0].toString() === rnOrderId.toString()) { rnTarget = rnRows[rni]; break; }
        }
      }
      if (!rnTarget) return ContentService.createTextOutput(JSON.stringify({result:"error",message:"找不到訂單"})).setMimeType(ContentService.MimeType.JSON);
      var rnResult = sendRefundEmail(ss, {
        orderId:       rnOrderId,
        eventName:     rnTarget[2] ? rnTarget[2].toString().trim() : "",
        custName:      rnTarget[8] ? rnTarget[8].toString() : "",
        custEmail:     rnTarget[10] ? rnTarget[10].toString().trim() : "",
        summary:       rnTarget[3] ? rnTarget[3].toString() : "",      // D欄：全部商品
        stockoutItems: rnTarget[19] ? rnTarget[19].toString() : "",    // T欄：缺貨商品
        stockoutAmt:   parseFloat(rnTarget[20]) || 0,                  // U欄：缺貨金額
        refundPhoto:   rnTarget[27] ? rnTarget[27].toString() : ""     // AB欄：退款截圖
      });
      if (rnResult) {
        // Q欄追加「已退款」標記
        var curQ = rnTarget[16] ? rnTarget[16].toString() : "";
        var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
        for (var rnFix = 1; rnFix < rnRows.length; rnFix++) {
          if (rnRows[rnFix][0].toString() === rnOrderId.toString()) {
            rnSheet.getRange(rnFix + 1, 17).setValue(curQ + " | 已退款:" + nowStr); // Q欄（記錄退款時間，不改 P 欄狀態）
            rnSheet.getRange(rnFix + 1, 29).setValue(nowStr);                        // AC欄：RefundedNotice
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({result:rnResult?"success":"error"})).setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 19：發送到貨通知 email ─────────────────────────────────────
    if (rowData.action === "sendArrivalNotification") {
      var anOrderId = rowData.orderId || "";
      var anSheet = ss.getSheetByName("訂單紀錄");
      var anTarget = null;
      if (anSheet) {
        var anRows = anSheet.getDataRange().getValues();
        for (var ai = 1; ai < anRows.length; ai++) {
          if (anRows[ai][0].toString() === anOrderId.toString()) { anTarget = anRows[ai]; break; }
        }
      }
      if (!anTarget) return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單" })).setMimeType(ContentService.MimeType.JSON);
      var anResult = sendArrivalEmail(ss, {
        orderId:        anOrderId,
        eventName:      anTarget[2] ? anTarget[2].toString().trim() : "",
        custName:       anTarget[8] ? anTarget[8].toString() : "",
        custEmail:      anTarget[10] ? anTarget[10].toString().trim() : "",
        summary:        anTarget[3] ? anTarget[3].toString() : "",
        shopUrl:        anTarget[18] ? anTarget[18].toString().trim() : "",  // S欄=賣貨便URL
        arrivalPhotoUrl:anTarget[23] ? anTarget[23].toString().trim() : "",  // X欄=到貨圖片
        shippingFee:    parseFloat(anTarget[5]) || 0,                        // F欄=運費
        weight:         parseFloat(anTarget[22]) || 0,                       // W欄=重量
        isResend:       rowData.isResend || false
      });
      if (anResult) {
        // 發送成功 → 改狀態為「已通知」
        var anSheet2 = ss.getSheetByName("訂單紀錄");
        if (anSheet2) {
          var anRows2 = anSheet2.getDataRange().getValues();
          for (var ai2 = 1; ai2 < anRows2.length; ai2++) {
            if (anRows2[ai2][0].toString() === anOrderId.toString()) {
              anSheet2.getRange(ai2 + 1, 16).setValue("已通知"); // P欄
              break;
            }
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ result: anResult ? "success" : "error" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 18：儲存賣貨便資料（W=重量, X=運費, Y=圖片）──────────────────────
    if (rowData.action === "saveSellshipData") {
      var ssSheet = ss.getSheetByName("訂單紀錄");
      if (!ssSheet) return ContentService.createTextOutput(JSON.stringify({ result: "error" })).setMimeType(ContentService.MimeType.JSON);
      var ssRows = ssSheet.getDataRange().getValues();
      var updates = rowData.updates || [];
      var saved = 0;
      for (var ui = 0; ui < updates.length; ui++) {
        var upd = updates[ui];
        for (var ur = 1; ur < ssRows.length; ur++) {
          if (ssRows[ur][0].toString() === upd.orderId.toString()) {
            if (upd.weight    !== undefined) ssSheet.getRange(ur + 1, 23).setValue(upd.weight);    // W欄
            if (upd.intlFee   !== undefined) ssSheet.getRange(ur + 1, 24).setValue(upd.intlFee);   // X欄
            if (upd.photoUrl  && upd.photoUrl !== "") ssSheet.getRange(ur + 1, 25).setValue(upd.photoUrl);  // Y欄
            saved++;
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ result: "success", saved: saved })).setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 18：賣貨便備貨資料儲存（F=運費, S=圖片URL, W=重量）──────────────
    if (rowData.action === "saveSfData") {
      var sfSheet = ss.getSheetByName("訂單紀錄");
      if (!sfSheet) return ContentService.createTextOutput(JSON.stringify({ result: "error" })).setMimeType(ContentService.MimeType.JSON);
      var sfRows = sfSheet.getDataRange().getValues();
      var sfUpdated = 0;
      var sfItems = rowData.items || []; // [{ orderId, url, weight, fee }]
      for (var si = 0; si < sfItems.length; si++) {
        var sfItem = sfItems[si];
        for (var sr = 1; sr < sfRows.length; sr++) {
          if (sfRows[sr][0].toString() === sfItem.orderId.toString()) {
            if (sfItem.url)    sfSheet.getRange(sr + 1, 24).setValue(sfItem.url);    // X欄：到貨圖片URL
            if (sfItem.fee)    sfSheet.getRange(sr + 1, 6).setValue(sfItem.fee);      // F欄：運費
            if (sfItem.weight) sfSheet.getRange(sr + 1, 23).setValue(sfItem.weight);  // W欄：重量
            sfUpdated++;
            // 即時檢查是否4欄齊全 → 改狀態為「已到貨」
            checkAndSetArrived(sfSheet, sr, sfSheet.getRange(sr + 1, 1, 1, 24).getValues()[0]);
            break;
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ result: "success", updated: sfUpdated }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 17：缺貨詢問通知 ─────────────────────────────────────
    if (rowData.action === "sendStockoutInquiry") {
      var siOrderId  = rowData.orderId || "";
      var siIsResend = rowData.isResend || false;
      var siOos      = rowData.outOfStockItems || [];
      var siIn       = rowData.inStockItems    || [];
      if (!siOrderId) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "缺少 orderId" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      // 從訂單紀錄取客人資料
      var siSheet = ss.getSheetByName("訂單紀錄");
      var siTarget = null;
      if (siSheet) {
        var siRows = siSheet.getDataRange().getValues();
        for (var si = 1; si < siRows.length; si++) {
          if (siRows[si][0].toString() === siOrderId.toString()) { siTarget = siRows[si]; break; }
        }
      }
      if (!siTarget) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      // isResend：從訂單紀錄的 T欄（缺貨商品）重建 siOos
      if (siIsResend) {
        var rStockout = siTarget[19] ? siTarget[19].toString() : "";
        var rSummary  = siTarget[3]  ? siTarget[3].toString()  : "";
        siOos = parseSummaryItems(rStockout).map(function(it){ return { name:it.name, sub:it.sub, qty:it.qty, price:0, photo:"" }; });
        siIn  = parseSummaryItems(rSummary).map(function(it){ return { name:it.name, sub:it.sub, qty:it.qty, price:0, photo:"" }; });
        // 過濾掉缺貨商品，只保留已購入
        var oosSet = {};
        siOos.forEach(function(it){ oosSet[it.name+"|"+(it.sub||"")] = true; });
        siIn = siIn.filter(function(it){ return !oosSet[it.name+"|"+(it.sub||"")]; });
      }
      var siPickup = siTarget[11] ? siTarget[11].toString() : "";
      var siPay    = siTarget[7]  ? siTarget[7].toString()  : "";
      var siIsTW   = siPickup.indexOf("賣貨便") !== -1 || siPay.indexOf("郵局") !== -1;
      var siEventName = siTarget[2] ? siTarget[2].toString().trim() : "";
      // 從 event sheet 補充商品價格（前端不傳價格）
      var siEventSheet2 = getSheetByEventName(ss, siEventName);
      var siPriceMap = {};
      if (siEventSheet2) {
        var siER = siEventSheet2.getDataRange().getValues();
        for (var ep2 = 1; ep2 < siER.length; ep2++) {
          var epN = siER[ep2][0] ? siER[ep2][0].toString().trim() : "";
          var epS = siER[ep2][1] ? siER[ep2][1].toString().trim() : "";
          var epY = parseFloat(siER[ep2][2]) || 0; // C欄=日幣
          var epH = parseFloat(siER[ep2][4]) || 0; // E欄=HKD
          var epT = parseFloat(siER[ep2][8]) || 0; // I欄=TWD
          if (epN) siPriceMap[epN + "|" + epS] = { yen: epY, hkd: epH, twd: epT };
        }
      }
      // 從 event sheet 取圖片 URL（F欄）
      var siPhotoMap = {};
      if (siEventSheet2) {
        for (var ep3 = 1; ep3 < siER.length; ep3++) {
          var epN3 = siER[ep3][0] ? siER[ep3][0].toString().trim() : "";
          var epF3 = siER[ep3][5] ? siER[ep3][5].toString().trim() : "";
          if (epN3 && epF3 && epF3.indexOf("http") === 0) siPhotoMap[epN3] = epF3;
        }
      }

      function enrichItems(items) {
        return items.map(function(it) {
          var pk = it.name + "|" + (it.sub || "");
          var pe = siPriceMap[pk] || siPriceMap[it.name + "|"] || {};
          // 如果精確比對失敗，嘗試去除 name 中的括號後比對
          if (!pe || (!pe.twd && !pe.hkd)) {
            var cleanN = it.name.replace(/\s*\([^)]+\)\s*$/, "").trim();
            pe = siPriceMap[cleanN + "|" + (it.sub || "")] || siPriceMap[cleanN + "|"] || pe || {};
            // 再嘗試模糊比對
            if (!pe.twd && !pe.hkd) {
              // 把 name 最後括號提取為 sub 再查
              var bM = it.name.match(/^(.*?)\s*[\(（]([^)）]+)[\)）]\s*$/);
              if (bM) {
                var bClean = bM[1].trim(), bSub = bM[2].trim();
                pe = siPriceMap[bClean + "|" + bSub] || siPriceMap[bClean + "|"] || pe || {};
              }
            }
          }
          // 最後 fallback：name 精確比對忽略 sub（用於商品名含款式但 sub 另存的情況）
          if (!pe || (!pe.twd && !pe.hkd)) {
            for (var fnk in siPriceMap) {
              if (fnk.split("|")[0] === it.name) { pe = siPriceMap[fnk]; break; }
            }
          }
          it.price = siIsTW ? (pe.twd || 0) : (pe.hkd || 0);
          // 直接從 event sheet 取圖片（最可靠）
          it.photo = siPhotoMap[it.name] || it.photo || "";
          return it;
        });
      }
      siOos = enrichItems(siOos);
      siIn  = enrichItems(siIn);

      var siResult = sendStockoutInquiryEmail(siIsTW, {
        isResend: siIsResend,
        orderId:          siOrderId,
        eventName:        siEventName,
        custName:         siTarget[8] ? siTarget[8].toString() : "",
        custEmail:        siTarget[10] ? siTarget[10].toString().trim() : "",
        outOfStockItems:  siOos,
        inStockItems:     siIn
      });

      if (siResult) { // 發信成功就更新，不論是否全部缺貨
        // ── 更新訂單：移除缺貨商品，重新計算金額 ──
        try {
          // 1. 重建 itemsSummary（只保留 inStockItems）
          var newSummaryParts = [];
          for (var ii = 0; ii < siIn.length; ii++) {
            var itm = siIn[ii];
            var part = itm.name;
            if (itm.sub) part += " (" + itm.sub + ")";
            // 保留原始 ( xN) 格式，避免再次解析時格式錯誤
            part += " ( x" + itm.qty + ")";
            newSummaryParts.push(part);
          }
          var newSummary = newSummaryParts.join(", ");

          // 2. 從 event sheet 取商品價格重算金額
          var siEventSheet = getSheetByEventName(ss, siEventName);
          var priceMapSI = {}; // { "名稱|款式": { hkd, twd } }
          if (siEventSheet) {
            var siERows = siEventSheet.getDataRange().getValues();
            for (var ep = 1; ep < siERows.length; ep++) {
              var epName = siERows[ep][0] ? siERows[ep][0].toString().trim() : "";
              var epSub  = siERows[ep][1] ? siERows[ep][1].toString().trim() : "";
              var epYEN  = parseFloat(siERows[ep][2]) || 0;
              var epHKD  = parseFloat(siERows[ep][4]) || 0;
              var epTWD  = parseFloat(siERows[ep][8]) || 0;
              if (epName) priceMapSI[epName + "|" + epSub] = { yen: epYEN, hkd: epHKD, twd: epTWD };
            }
          }

          // ── 計算缺貨商品的金額（寫入 T/U/V 欄）──
          var oosSummaryParts = [];
          var oosTotal = 0;
          var oosYenTotal = 0;
          for (var io = 0; io < siOos.length; io++) {
            var oosItm = siOos[io];
            var oosPart = oosItm.name;
            if (oosItm.sub) oosPart += " (" + oosItm.sub + ")";
            oosPart += " ( x" + oosItm.qty + ")";
            oosSummaryParts.push(oosPart);
            var oosPrice = findPrice(oosItm.name, oosItm.sub);
            if (oosPrice) {
              var oosUnit = siIsTW ? oosPrice.twd : oosPrice.hkd;
              oosTotal    += oosUnit * (oosItm.qty || 1);
              oosYenTotal += (oosPrice.yen || 0) * (oosItm.qty || 1);
            }
          }
          var oosSummary = oosSummaryParts.join("，");

          // ── 計算保留商品的新金額（更新 D/E/G 欄）──
          // 模糊查找價格（處理 name 包含款式的情況）
          function findPrice(iName, iSub) {
            // 1. 精確比對 name|sub
            var exact = siPriceMap[iName + "|" + (iSub || "")] || siPriceMap[iName + "|"];
            if (exact) return exact;
            // 2. 把 name 最後的括號提取為 sub 再比對
            //    e.g. "NIRACHAN mini NUI (綺羅キラー)" → name="NIRACHAN mini NUI", sub="綺羅キラー"
            var bracketM = iName.match(/^(.*?)\s*[\(（]([^)）]+)[\)）]\s*$/);
            if (bracketM) {
              var cleanName   = bracketM[1].trim();
              var extractedSub = bracketM[2].trim();
              // 2a. cleanName | extractedSub
              var v1 = siPriceMap[cleanName + "|" + extractedSub];
              if (v1) return v1;
              // 2b. cleanName | iSub（原本的 sub）
              if (iSub) {
                var v2 = siPriceMap[cleanName + "|" + iSub];
                if (v2) return v2;
              }
              // 2c. cleanName |（空）
              var v3 = siPriceMap[cleanName + "|"];
              if (v3) return v3;
            }
            // 最後 fallback：只用 name 精確比對（忽略 sub）
            for (var ffk in siPriceMap) {
              if (ffk.split("|")[0] === iName) return siPriceMap[ffk];
            }
            return null;
          }

          var newTotal = 0;
          var newYenTotal = 0;
          for (var ij = 0; ij < siIn.length; ij++) {
            var itm2 = siIn[ij];
            var priceEntry = findPrice(itm2.name, itm2.sub);
            if (priceEntry) {
              var unitPrice = siIsTW ? priceEntry.twd : priceEntry.hkd;
              newTotal += unitPrice * (itm2.qty || 1);
              newYenTotal += (priceEntry.yen || 0) * (itm2.qty || 1);
            }
          }

          // 加運費（F欄）
          var shippingFee = parseFloat(siTarget[5]) || 0;
          newTotal += shippingFee;

          // 3. 找訂單行並更新
          var siAllRows = siSheet.getDataRange().getValues();
          for (var sr2 = 1; sr2 < siAllRows.length; sr2++) {
            if (siAllRows[sr2][0].toString() === siOrderId.toString()) {
              siSheet.getRange(sr2 + 1, 4).setValue(newSummary);    // D欄：商品明細（只含已購入）
              siSheet.getRange(sr2 + 1, 5).setValue(newTotal);       // E欄：港幣/台幣總額（只含已購入）
              siSheet.getRange(sr2 + 1, 7).setValue(newYenTotal);    // G欄：日幣總額（只含已購入）
              siSheet.getRange(sr2 + 1, 20).setValue(oosSummary);    // T欄：Stockout Items
              siSheet.getRange(sr2 + 1, 21).setValue(oosTotal);      // U欄：Stockout Amount
              siSheet.getRange(sr2 + 1, 22).setValue(oosYenTotal);   // V欄：Stockout JPY
              // Q欄記錄發送時間
              var siNowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
              siSheet.getRange(sr2 + 1, 17).setValue("已寄出缺貨通知:" + siNowStr);

              // 如果全部商品缺貨（inStockItems 為空），將 P欄改為「已取消」
              if (siIn.length === 0) {
                siSheet.getRange(sr2 + 1, 16).setValue("已取消");
                Logger.log("缺貨全部取消訂單 " + siOrderId);
              }

              Logger.log("缺貨更新訂單 " + siOrderId + "：剩餘=" + newSummary + "，缺貨=" + oosSummary);
              break;
            }
          }
        } catch(updateErr) {
          Logger.log("缺貨更新訂單失敗: " + updateErr.toString());
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ result: siResult ? "success" : "error" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 16：發送購物清單 email ────────────────────────────────
    if (rowData.action === "sendShoppingList") {
      var evName  = rowData.eventName || "";
      var slItems = rowData.items || [];
      if (!evName || !slItems.length) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "缺少參數" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var slResult = sendShoppingList(ss, evName, slItems);
      return ContentService.createTextOutput(JSON.stringify(slResult))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 14：儲存港幣總成本到盈利紀錄 ─────────────────────────
    if (rowData.action === "saveProfitCost") {
      var evName    = rowData.eventName ? rowData.eventName.toString().trim() : "";
      var totalCost = parseFloat(rowData.totalCost) || 0;
      var twdRate   = parseFloat(rowData.twdRate) || 3.97;
      if (!evName) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "缺少 eventName" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      saveProfitCost(ss, evName, totalCost, twdRate);
      return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 13：儲存結算數量到購貨紀錄 F欄 ──────────────────────
    if (rowData.action === "saveSettled") {
      var evName = rowData.eventName ? rowData.eventName.toString().trim() : "";
      var sItems = rowData.items; // [{ name, sub, settled }]
      if (!evName || !Array.isArray(sItems)) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "缺少參數" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      // 優先使用 per-event sheet
      var sPerEv = getEventPurchaseSheet(ss, evName, false);
      var sSheet = sPerEv || getPurchaseSheet(ss);
      var sUsePerEv = !!sPerEv;
      if (!sSheet) {
        sSheet = ss.insertSheet("購貨紀錄");
        sSheet.appendRow(["Event名稱", "商品名稱", "款式", "日幣單價", "已付款數量", "未付款數量", "已購買數量", "已結算數量"]);
        sUsePerEv = false;
      }
      var sRows = sSheet.getDataRange().getValues();
      for (var si = 0; si < sItems.length; si++) {
        var sItem = sItems[si];
        var found = false;
        var sRows2 = sSheet.getDataRange().getValues();
        for (var sr = 1; sr < sRows2.length; sr++) {
          var srName = sUsePerEv
            ? (sRows2[sr][0] ? sRows2[sr][0].toString().trim() : "")
            : (sRows2[sr][1] ? sRows2[sr][1].toString().trim() : "");
          var srSub = sUsePerEv
            ? (sRows2[sr][1] ? sRows2[sr][1].toString().trim() : "")
            : (sRows2[sr][2] ? sRows2[sr][2].toString().trim() : "");
          if (!sUsePerEv) {
            var srEv = sRows2[sr][0] ? sRows2[sr][0].toString().trim() : "";
            if (srEv !== evName) continue;
          }
          if (srName === sItem.name && srSub === (sItem.sub || '')) {
            var settledCol = sUsePerEv ? 7 : 8; // per-event G=col7; 舊版 H=col8
            var curSettled = Number(sRows2[sr][settledCol - 1]) || 0;
            sSheet.getRange(sr + 1, settledCol).setValue(curSettled + sItem.settled); // 累加
            found = true;
            break;
          }
        }
        if (!found) {
          // 新增一行
          sSheet.appendRow([evName, sItem.name, sItem.sub, 0, 0, sItem.settled]);
          sRows.push([evName, sItem.name, sItem.sub, 0, 0, sItem.settled]);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ result: "success", saved: sItems.length }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 12：更新訂單欄位 ──────────────────────────────────────
    if (rowData.action === "updateOrder") {
      var uSheet = ss.getSheetByName("訂單紀錄");
      if (!uSheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單紀錄" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var uRows = uSheet.getDataRange().getValues();
      var uTargetRow = -1;
      for (var r = 1; r < uRows.length; r++) {
        if (uRows[r][0].toString() === rowData.orderId.toString()) { uTargetRow = r; break; }
      }
      if (uTargetRow === -1) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單編號" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var rowNum = uTargetRow + 1; // 1-indexed
      var fields = rowData.fields || {};
      // 欄位對照表（欄位名 → Sheet 欄號，1-indexed）
      var colMap = {
        itemsSummary:   4,  // D
        totalAmount:    5,  // E
        shippingFee:    6,  // F
        totalJPYAmount: 7,  // G
        paymentMethod:  8,  // H
        recipientName:  9,  // I
        recipientPhone: 10, // J
        recipientEmail: 11, // K
        pickupType:     12, // L
        pickupCode:     13, // M
        pickupName:     14, // N
        pickupAddress:  15, // O
        status:         16, // P
        remark:         17, // Q
      };
      for (var field in fields) {
        if (colMap[field]) {
          uSheet.getRange(rowNum, colMap[field]).setValue(fields[field]);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 15：手動發送催款信 ──────────────────────────────────────
    if (rowData.action === "sendManualReminder") {
      var rSheet = ss.getSheetByName("訂單紀錄");
      if (!rSheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單紀錄" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var rRows = rSheet.getDataRange().getValues();
      var rTarget = -1;
      for (var r = 1; r < rRows.length; r++) {
        if (rRows[r][0].toString() === rowData.orderId.toString()) { rTarget = r; break; }
      }
      if (rTarget === -1) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var rRow = rRows[rTarget];
      var rPickup = rRow[11] ? rRow[11].toString() : "";
      var rPay    = rRow[7]  ? rRow[7].toString()  : "";
      var isTW    = rPickup.indexOf("賣貨便") !== -1 || rPay.indexOf("郵局") !== -1;
      var sent = sendReminderEmail(isTW, {
        orderId:     rRow[0].toString(),
        orderDate:   rRow[1],
        eventName:   rRow[2] ? rRow[2].toString().trim() : "",
        totalAmount: rRow[4] || 0,
        payMethod:   rPay,
        custName:    rRow[8] ? rRow[8].toString() : "",
        custPhone:   rRow[9] ? rRow[9].toString() : "",
        custEmail:   rRow[10] ? rRow[10].toString().trim() : "",
        summary:     rRow[3] ? rRow[3].toString() : ""
      });
      // 更新 Q欄標記
      if (sent) rSheet.getRange(rTarget + 1, 17).setValue("已催款");
      return ContentService.createTextOutput(JSON.stringify({ result: sent ? "success" : "error" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 11：重新發送待付款確認信 ──────────────────────────────
    if (rowData.action === "resendPendingEmail") {
      var oSheet = ss.getSheetByName("訂單紀錄");
      if (!oSheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單紀錄" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var oRows = oSheet.getDataRange().getValues();
      var oTargetRow = -1;
      for (var r = 1; r < oRows.length; r++) {
        if (oRows[r][0].toString() === rowData.orderId.toString()) { oTargetRow = r; break; }
      }
      if (oTargetRow === -1) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單編號" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var oRow = oRows[oTargetRow];
      // 重組成跟原本下單一樣的 rowData 格式，然後呼叫發信邏輯
      var reRowData = {
        region:         oRow[11] && oRow[11].toString().indexOf("賣貨便") !== -1 ? "tw" :
                        oRow[7]  && oRow[7].toString().indexOf("郵局") !== -1    ? "tw" : "hk",
        eventSheetName: oRow[2]  ? oRow[2].toString().trim() : "",
        itemsSummary:   oRow[3]  ? oRow[3].toString() : "",
        totalAmount:    oRow[4]  || 0,
        shippingFee:    oRow[5]  || 0,
        totalJPYAmount: oRow[6]  || 0,
        paymentMethod:  oRow[7]  ? oRow[7].toString() : "",
        recipientName:  oRow[8]  ? oRow[8].toString() : "",
        recipientPhone: oRow[9]  ? oRow[9].toString() : "",
        recipientEmail: oRow[10] ? oRow[10].toString().trim() : "",
        pickupType:     oRow[11] ? oRow[11].toString() : "",
        pickupCode:     oRow[12] ? oRow[12].toString() : "",
        pickupName:     oRow[13] ? oRow[13].toString() : "",
        pickupAddress:  oRow[14] ? oRow[14].toString() : ""
      };
      var reOrderId = oRow[0].toString();
      var reRegion  = reRowData.region;
      var reFinalEventName = reRowData.eventSheetName || "現場快閃代購";
      var reTargetEmail    = reRowData.recipientEmail;

      if (!reTargetEmail || reTargetEmail.indexOf("@") === -1) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "無效的客戶 Email" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }

      // 讀取商品資料（用於圖片比對）
      var reEventSheet = ss.getSheetByName(reFinalEventName);
      var reBackupProducts = [];
      if (reEventSheet) reBackupProducts = reEventSheet.getDataRange().getValues();

      // 解析 cartItems（從 itemsSummary）
      var reParsedItems = parseSummaryItems(reRowData.itemsSummary);
      var reCartItems = [];
      for (var ri = 0; ri < reParsedItems.length; ri++) {
        var rItem = reParsedItems[ri];
        // 從商品分頁查圖片和價格
        var rPhoto = "", rPrice = 0, rTwd = 0, rYen = 0;
        var cleanRName = rItem.name.toLowerCase().replace(/[^a-z0-9一-龥]/g, "");
        for (var rp = 1; rp < reBackupProducts.length; rp++) {
          var rpName = reBackupProducts[rp][0] ? reBackupProducts[rp][0].toString().trim() : "";
          var cleanRpName = rpName.toLowerCase().replace(/[^a-z0-9一-龥]/g, "");
          if (cleanRName && cleanRpName && (cleanRName.indexOf(cleanRpName) !== -1 || cleanRpName.indexOf(cleanRName) !== -1)) {
            rPhoto = reBackupProducts[rp][5] ? reBackupProducts[rp][5].toString().trim() : "";
            rPrice = parseFloat(reBackupProducts[rp][4]) || 0;
            rTwd   = parseFloat(reBackupProducts[rp][8]) || 0;
            rYen   = parseFloat(reBackupProducts[rp][2]) || 0;
            break;
          }
        }
        reCartItems.push({ name: rItem.name, sub: rItem.sub, qty: rItem.qty,
                           price: rPrice, twd: rTwd, yen: rYen, photo: rPhoto });
      }
      reRowData.cartItems = reCartItems;

      // 重新發送待付款 email（HK/TW 分流，邏輯與原本完全一致）
      if (reRegion === "tw") {
        // ── TW 待付款 email ──
        var twPhoneLast3 = reRowData.recipientPhone ? reRowData.recipientPhone.toString().slice(-3) : "";

        // 計算付款期限
        var reDeadlineMap = syncAndGetDeadlines(ss);
        var reEventDlStr = reDeadlineMap[reFinalEventName] || "";
        var re72h = new Date(new Date().getTime() + 72 * 60 * 60 * 1000);
        var rePayDeadlineTextTW = "請於20分鐘內完成匯款";
        if (reEventDlStr) {
          var reEventDl = new Date(reEventDlStr.replace(/\//g, "-"));
          if (!isNaN(reEventDl) && reEventDl < re72h) {
            var dlM = reEventDl.getMonth()+1, dlD = reEventDl.getDate();
            var dlH = reEventDl.getHours(), dlMin = reEventDl.getMinutes() < 10 ? "0"+reEventDl.getMinutes() : reEventDl.getMinutes();
            rePayDeadlineTextTW = "請於" + dlM + "月" + dlD + "日 " + dlH + ":" + dlMin + " 前完成匯款";
          }
        }

        // 建立 TW 商品表格
        var reTwTableRows = "";
        var reTwPhotoMap = {};
        for (var rbp = 1; rbp < reBackupProducts.length; rbp++) {
          var rbpN = reBackupProducts[rbp][0] ? reBackupProducts[rbp][0].toString().trim() : "";
          var rbpP = reBackupProducts[rbp][5] ? reBackupProducts[rbp][5].toString().trim() : "";
          if (rbpN) reTwPhotoMap[rbpN] = rbpP;
        }
        var fallbackImg = "data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%221%22%20height%3D%221%22%3E%3Crect%20width%3D%221%22%20height%3D%221%22%20fill%3D%22%23111%22%2F%3E%3C%2Fsvg%3E";
        for (var ti = 0; ti < reCartItems.length; ti++) {
          var tItem = reCartItems[ti];
          var tDn = tItem.name + (tItem.sub ? " (" + tItem.sub + ")" : "");
          var tQty = tItem.qty || 1;
          var tPrice = parseFloat(tItem.twd) || 0;
          var tSubtotal = tPrice * tQty;
          var tImgUrl = (tItem.photo && tItem.photo.indexOf("http") === 0) ? tItem.photo : fallbackImg;
          reTwTableRows += "<tr style=\"border-bottom:1px solid #eeeeee;\">" +
            "<td style=\"padding:12px 10px 12px 15px; font-size:14px; color:#333;\">" +
              "<table style=\"width:100%; border-collapse:collapse; border:none;\"><tr>" +
                "<td style=\"padding:0; width:50px; vertical-align:middle;\">" +
                  "<img src=\"" + tImgUrl + "\" style=\"width:50px; height:50px; object-fit:cover; border-radius:4px; display:block;\" />" +
                "</td>" +
                "<td style=\"padding:0 0 0 12px; vertical-align:middle;\">" + tDn + "</td>" +
              "</tr></table>" +
            "</td>" +
            "<td style=\"padding:12px 10px; font-size:14px; color:#555; text-align:center;\">" + tQty + "</td>" +
            "<td style=\"padding:12px 15px; font-size:14px; color:#2980b9; text-align:right; font-weight:500;\">" +
              (tSubtotal > 0 ? "NT$ " + tSubtotal : "—") +
            "</td></tr>";
        }

        var reTwSubject = "[已收到訂單] 代購 " + reFinalEventName + " — 886tw.81jp 台灣";
        var reTwBody = "<div style=\"font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width:600px; margin:0 auto; color:#333; line-height:1.7; padding:20px; background:#f7f6f3;\">" +
          "<h2 style=\"font-size:18px; font-weight:500; border-bottom:1px solid #e5e5e5; padding-bottom:10px; color:#1a1a1a;\">代購 " + reFinalEventName + "</h2>" +
          "<p style=\"font-size:14px;\">您好！我們已收到您的訂單 📋</p><p style=\\\"font-size:13px; background:#fff0f0; border-left:4px solid #e53e3e; padding:11px 14px; border-radius:4px; margin:10px 0; color:#c53030; font-weight:600; line-height:1.7;\\\">⚠️ 重要：我們會在收到款項後才會進行採購，目前訂單狀態為【待付款】，請完成匯款以確保成功訂購。</p><p style=\\\"font-size:13px; background:#fff8e1; border-left:4px solid #f59e0b; padding:11px 14px; border-radius:4px; margin:10px 0; color:#78350f; line-height:1.7;\\\">⏰ 請於 <strong>20分鐘內</strong> 完成匯款，逾時訂單將自動取消，如果還需要代購請在代購活動截止前再次提交。如遇缺貨，將以付款先後次序分配商品。</p><p style=\\\"font-size:13px; color:#888; margin-top:6px; line-height:1.7;\">如遇缺貨，將以付款順序分配商品。</p>" +
          "<p style=\"font-size:14px; background:#f2f2f2; padding:8px 12px; border-radius:4px; display:inline-block;\"><strong>訂單編號：</strong>" + reOrderId + "</p>" +
          "<div style=\"margin:35px 0; padding:0 10px;\">" +"<div style=\"position:relative; display:flex; justify-content:space-between; align-items:center; width:100%;\">" +"<div style=\"position:absolute; top:5px; left:0; right:0; height:1px; background-color:#e5e5e5; z-index:1;\"></div>" +"<div style=\"position:relative; z-index:2; text-align:center; width:25%;\">" +"<div style=\"width:11px; height:11px; background-color:#2980b9; border-radius:50%; margin:0 auto; border:2px solid #ffffff; box-shadow:0 0 0 1px #2980b9;\"></div>" +"<div style=\"font-size:12px; font-weight:bold; color:#2980b9; margin-top:8px;\">已收到訂單</div></div>" +"<div style=\"position:relative; z-index:2; text-align:center; width:25%;\">" +"<div style=\"width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;\"></div>" +"<div style=\"font-size:11px; color:#888888; margin-top:8px;\">已付款</div></div>" +"<div style=\"position:relative; z-index:2; text-align:center; width:25%;\">" +"<div style=\"width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;\"></div>" +"<div style=\"font-size:11px; color:#888888; margin-top:8px;\">已到台</div></div>" +"<div style=\"position:relative; z-index:2; text-align:center; width:25%;\">" +"<div style=\"width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;\"></div>" +"<div style=\"font-size:11px; color:#888888; margin-top:8px;\">已完成取貨</div></div>" +"</div></div>" +
          "<h3 style=\"font-size:14px; font-weight:bold; margin-top:25px; margin-bottom:10px; color:#444;\">📋 訂購商品明細</h3>" +
          "<table style=\"width:100%; border-collapse:collapse; background:#fff; border:1px solid #eef0ed; border-radius:6px; overflow:hidden;\">" +
            "<thead><tr style=\"background-color:#f0f7ff; border-bottom:1px solid #eef0ed;\">" +
              "<th style=\"padding:10px 15px; text-align:left; font-size:13px; color:#666; font-weight:500;\">商品名稱與款式</th>" +
              "<th style=\"padding:10px; text-align:center; font-size:13px; color:#666; font-weight:500; width:60px;\">數量</th>" +
              "<th style=\"padding:10px 15px; text-align:right; font-size:13px; color:#666; font-weight:500; width:90px;\">小計</th>" +
            "</tr></thead>" +
            "<tbody>" + reTwTableRows + "</tbody>" +
          "</table>" +
          "<p style=\"text-align:right; font-size:16px; font-weight:bold; margin-top:15px; color:#1a1a1a;\">商品小計：<span style=\"font-size:20px; color:#2980b9;\">NT$ " + reRowData.totalAmount + "</span></p>" +
          "<p style=\"text-align:right; font-size:12px; color:#888; margin-top:4px;\">＋ 國際運費（NT$20／50g）＋ NT$38 賣貨便運費（出貨時另行通知）</p>" +
          "<div style=\"background-color:#f0f7ff; border-left:4px solid #2980b9; padding:15px; margin:20px 0; border-radius:4px;\">" +
            "<p style=\"margin:0; font-weight:bold; color:#1565c0; font-size:15px;\">💰 郵局銀行匯款資訊</p>" +
            "<p style=\\\"margin:6px 0 10px; font-size:15px; color:#1a237e;\\\">匯款金額：<strong style=\\\"font-size:20px;\\\">NT$ " + reRowData.totalAmount + "</strong></p>" +
            "<p style=\"margin:10px 0 4px; font-size:14px; color:#333;\">銀行：<strong>中華郵政（郵局）</strong></p>" +
            "<p style=\"margin:4px 0; font-size:14px; color:#333;\">局號／帳號：<strong>0041860-0025565</strong></p>" +
            "<p style=\"margin:4px 0; font-size:14px; color:#333;\">戶名：<strong>周◯恩</strong></p>" +
            "<p style=\"margin:12px 0 6px; font-size:14px; font-weight:700; color:#1565c0;\">" +
              "⚠️ " + rePayDeadlineTextTW + "，匯款時請於備註欄填寫您的<strong>手機末3碼</strong>：" +
            "</p>" +
            "<span style=\"display:inline-block; background:#fff; border:2px dashed #2980b9; padding:6px 18px; border-radius:6px; font-family:monospace; font-size:22px; font-weight:bold; color:#2980b9; margin:4px 0 10px; letter-spacing:6px;\">" + twPhoneLast3 + "</span>" +
            "<p style=\"margin:8px 0 0; font-size:12px; color:#888; line-height:1.7;\">匯款後如24小時內沒有收到已付款通知，請主動回覆電郵或<a href=\"https://www.threads.net/@852hk.81jp\" style=\"color:#2980b9; font-weight:600;\">私訊我們</a>。</p>" +
          "</div>" +
          "<div style=\"margin-top:30px; padding-top:15px; border-top:1px dashed #ddd; font-size:12px; color:#777; line-height:1.8;\">" +
            "<p>✅ 我們收到匯款後才會進行採購，並發送付款確認 Email 通知您。</p>" +
            "<p>如有任何問題，歡迎直接回覆此 Email 或於 Threads 私訊我們查詢，謝謝！</p>" +
          "</div>" +
          "<div style=\"margin-top:30px; font-size:13px; color:#333;\">" +
            "<strong>852hk.81jp</strong><br>" +
            "Threads: <a href=\"https://www.threads.net/@852hk.81jp\" style=\"color:#2980b9;\">https://www.threads.net/@852hk.81jp</a>" +
          "</div>" +
          "</div>";

        MailApp.sendEmail({ to: reTargetEmail, subject: reTwSubject,
          body: "感謝您的訂購！訂單編號：" + reOrderId, htmlBody: reTwBody });

      } else {
        // ── HK 待付款 email ──
        var reDeadlineMapHK = syncAndGetDeadlines(ss);
        var reEventDlStrHK  = reDeadlineMapHK[reFinalEventName] || "";
        var re72hHK = new Date(new Date().getTime() + 72 * 60 * 60 * 1000);
        var rePayDeadlineTextHK = "請於20分鐘內完成付款";
        if (reEventDlStrHK) {
          var reEventDlHK = new Date(reEventDlStrHK.replace(/\//g, "-"));
          if (!isNaN(reEventDlHK) && reEventDlHK < re72hHK) {
            var dlMH = reEventDlHK.getMonth()+1, dlDH = reEventDlHK.getDate();
            var dlHH = reEventDlHK.getHours(), dlMinH = reEventDlHK.getMinutes() < 10 ? "0"+reEventDlHK.getMinutes() : reEventDlHK.getMinutes();
            rePayDeadlineTextHK = "請於" + dlMH + "月" + dlDH + "日 " + dlHH + ":" + dlMinH + " 前完成付款";
          }
        }

        var reCopyCodeHtml = "<span style=\"display:inline-block; background-color:#fff; border:1px dashed #ccc; padding:4px 10px; border-radius:4px; font-family:monospace; font-size:15px; color:#111; font-weight:bold;\">" + reOrderId + "</span>";
        var rePaymentDetailsHtml = "";
        var rePayMethodUpper = reRowData.paymentMethod ? reRowData.paymentMethod.toString().toUpperCase() : "";

        if (rePayMethodUpper.indexOf("PAYME") !== -1) {
          rePaymentDetailsHtml =
            "<div style=\"background-color:#fff5f5; border-left:4px solid #ff4d4d; padding:15px; margin:15px 0; border-radius:4px;\">" +
              "<p style=\"margin:0; font-weight:bold; color:#d93838; font-size:15px;\">💰 PayMe 付款資料：</p>" +
              "<p style=\\\"margin:6px 0 10px; font-size:15px; color:#b71c1c;\\\">匯款金額：<strong style=\\\"font-size:20px;\\\">HK$ " + reRowData.totalAmount + "</strong></p>" +
              "<p style=\"margin:8px 0; font-size:14px; color:#333;\"><a href=\"https://payme.hsbc/miru\" style=\"color:#ff4d4d; font-weight:bold;\">一按即 PayMe！ https://payme.hsbc/miru</a></p>" +
              "<p style=\"margin:8px 0 0; font-size:13px; color:#666;\">⚠️ " + rePayDeadlineTextHK + "，付款時請於【備註】輸入訂單編號：" + reCopyCodeHtml + "</p>" +
            "</div>";
        } else if (rePayMethodUpper.indexOf("FPS") !== -1 || rePayMethodUpper.indexOf("轉數快") !== -1) {
          rePaymentDetailsHtml =
            "<div style=\"background-color:#f0f7ff; border-left:4px solid #0066cc; padding:15px; margin:15px 0; border-radius:4px;\">" +
              "<p style=\"margin:0; font-weight:bold; color:#0052a3; font-size:15px;\">💰 轉數快 (FPS) 付款資料：</p>" +
              "<p style=\"margin:8px 0; font-size:14px; color:#333;\">轉數快快速支付號碼：<strong>8890873</strong><br>帳戶持有人：<strong>Chow W. Y.</strong></p>" +
              "<p style=\\\"margin:6px 0 10px; font-size:15px; color:#1a237e;\\\">匯款金額：<strong style=\\\"font-size:20px;\\\">HK$ " + reRowData.totalAmount + "</strong></p>" +
              "<p style=\"margin:8px 0 0; font-size:13px; color:#666;\">⚠️ " + rePayDeadlineTextHK + "，付款時請於【備註】輸入訂單編號：" + reCopyCodeHtml + "</p>" +
            "</div>";
        }

        // 建立 HK 商品表格
        var reHkTableRows = "";
        for (var hi = 0; hi < reCartItems.length; hi++) {
          var hItem = reCartItems[hi];
          var hDn = hItem.name + (hItem.sub ? " (" + hItem.sub + ")" : "");
          var hQty = hItem.qty || 1;
          var hPrice = parseFloat(hItem.price) || 0;
          var hSubtotal = hPrice * hQty;
          var hImgUrl = (hItem.photo && hItem.photo.indexOf("http") === 0) ? hItem.photo : fallbackImg;
          reHkTableRows += "<tr style=\"border-bottom:1px solid #eee;\">" +
            "<td style=\"padding:12px 10px 12px 15px; font-size:14px; color:#333;\">" +
              "<table style=\"width:100%; border-collapse:collapse; border:none;\"><tr>" +
                "<td style=\"padding:0; width:50px; vertical-align:middle;\">" +
                  "<img src=\"" + hImgUrl + "\" style=\"width:50px; height:50px; object-fit:cover; border-radius:4px; display:block;\" />" +
                "</td>" +
                "<td style=\"padding:0 0 0 12px; vertical-align:middle;\">" + hDn + "</td>" +
              "</tr></table>" +
            "</td>" +
            "<td style=\"padding:12px 10px; font-size:14px; color:#555; text-align:center;\">" + hQty + "</td>" +
            "<td style=\"padding:12px 15px; font-size:14px; color:#111; text-align:right; font-weight:500;\">" +
              "HK$ " + (hSubtotal > 0 ? hSubtotal : "—") +
            "</td></tr>";
        }

        var reHkSubject = "[已收到訂單] " + reFinalEventName + " - 852hk.81jp";
        var reHkBody = "<div style=\"font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width:600px; margin:0 auto; color:#333; line-height:1.6; padding:20px; background:#fbfbfa;\">" +
          "<h2 style=\"font-size:18px; font-weight:500; border-bottom:1px solid #e5e5e5; padding-bottom:10px; color:#111;\">代購 " + reFinalEventName + "</h2>" +
          "<p style=\"font-size:14px;\">你好，我哋已收到你嘅訂單 📋</p><p style=\\\"font-size:13px; background:#fff0f0; border-left:4px solid #e53e3e; padding:11px 14px; border-radius:4px; margin:10px 0; color:#c53030; font-weight:600; line-height:1.7;\\\">⚠️ 重要：我哋係收到款項後先至會購買商品，現時訂單狀態為【待付款】，請完成付款以確保成功訂購。</p><p style=\"font-size:13px; color:#888; margin-top:8px; line-height:1.7;\">📌 我哋係<strong>收到款項後</strong>先至會購買商品，現時訂單狀態為【待付款】。<br>⏰ 請於 <strong>20分鐘內</strong> 完成付款，逾時訂單將自動取消。如遇缺貨，將以付款先後次序分配商品。</p><p style=\\\"font-size:13px; color:#888; margin-top:6px; line-height:1.7;\">如遇缺貨，將以付款先後次序分配商品。</p>" +
          "<p style=\"font-size:14px; background:#f2f2f2; padding:8px 12px; border-radius:4px; display:inline-block;\"><strong>訂單編號：</strong>" + reOrderId + "</p>" +
          "<div style=\"margin:35px 0; padding:0 10px;\">" +"<div style=\"position:relative; display:flex; justify-content:space-between; align-items:center; width:100%;\">" +"<div style=\"position:absolute; top:5px; left:0; right:0; height:1px; background-color:#e5e5e5; z-index:1;\"></div>" +"<div style=\"position:relative; z-index:2; text-align:center; width:25%;\">" +"<div style=\"width:11px; height:11px; background-color:#111111; border-radius:50%; margin:0 auto; border:2px solid #ffffff; box-shadow:0 0 0 1px #111111;\"></div>" +"<div style=\"font-size:12px; font-weight:bold; color:#111111; margin-top:8px;\">已收到訂單</div></div>" +"<div style=\"position:relative; z-index:2; text-align:center; width:25%;\">" +"<div style=\"width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;\"></div>" +"<div style=\"font-size:11px; color:#888888; margin-top:8px;\">已付款</div></div>" +"<div style=\"position:relative; z-index:2; text-align:center; width:25%;\">" +"<div style=\"width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;\"></div>" +"<div style=\"font-size:11px; color:#888888; margin-top:8px;\">已付款</div></div>" +"<div style=\"position:relative; z-index:2; text-align:center; width:25%;\">" +"<div style=\"width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;\"></div>" +"<div style=\"font-size:11px; color:#888888; margin-top:8px;\">已寄出</div></div>" +"</div></div>" +
          "<h3 style=\"font-size:14px; font-weight:bold; margin-top:25px; margin-bottom:10px; color:#444;\">📋 訂購商品明細</h3>" +
          "<table style=\"width:100%; border-collapse:collapse; background:#fff; border:1px solid #eef0ed; border-radius:6px; overflow:hidden;\">" +
            "<thead><tr style=\"background-color:#fafbfa; border-bottom:1px solid #eef0ed;\">" +
              "<th style=\"padding:10px 15px; text-align:left; font-size:13px; color:#666;\">商品名稱與款式</th>" +
              "<th style=\"padding:10px; text-align:center; font-size:13px; color:#666; width:60px;\">數量</th>" +
              "<th style=\"padding:10px 15px; text-align:right; font-size:13px; color:#666; width:80px;\">小計</th>" +
            "</tr></thead>" +
            "<tbody>" + reHkTableRows + "</tbody>" +
          "</table>" +
          "<p style=\"text-align:right; font-size:14px; margin-top:15px; color:#666;\">運費：<span>HK$ " + reRowData.shippingFee + "</span></p>" +
          "<p style=\"text-align:right; font-size:16px; font-weight:bold; color:#111;\">總金額（含郵費）：<span style=\"font-size:20px; color:#ff4d4d;\">HK$ " + reRowData.totalAmount + "</span></p>" +
          "<h3 style=\"font-size:14px; font-weight:bold; margin-top:25px; margin-bottom:5px; color:#444;\">📍 易寄取收件資料</h3>" +
          "<div style=\"background:#fff; border:1px solid #eef0ed; padding:15px; border-radius:6px; font-size:13px; color:#555; line-height:1.8;\">" +
            "<strong>收件人姓名：</strong>" + reRowData.recipientName + "<br>" +
            "<strong>聯絡電話：</strong>" + reRowData.recipientPhone + "<br>" +
            "<strong>自提點名稱：</strong>" + reRowData.pickupName + "<br>" +
            "<strong>自提點地址：</strong>" + reRowData.pickupAddress +
          "</div>" +
          "<p style=\"font-size:14px; margin-top:25px; font-weight:bold;\">💡 如資料無誤，請使用下方資訊完成付款：</p>" +
          rePaymentDetailsHtml +
          "<div style=\"margin-top:30px; padding-top:15px; border-top:1px dashed #ddd; font-size:12px; color:#777;\">" +
            "<p>✅ 收到款項後，我哋會購買商品並發送付款確認 Email 通知你。</p>" +
            "<p>如有任何問題，歡迎直接回覆此 Email 或到 Threads inbox 我哋查詢，謝謝！</p>" +
          "</div>" +
          "</div>";

        MailApp.sendEmail({ to: reTargetEmail, subject: reHkSubject,
          body: "感謝您的訂購！訂單編號: " + reOrderId, htmlBody: reHkBody });
      }

      return ContentService.createTextOutput(JSON.stringify({ result: "success", orderId: reOrderId }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 9：發送取消訂單 email（A=未付款 / C=客人要求）──────────
    if (rowData.action === "sendCancelEmail") {
      var result = sendCancelEmail(ss, rowData.orderId, rowData.reason);
      return ContentService.createTextOutput(JSON.stringify(result))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 10：發送缺貨取消 email ──────────────────────────────────
    if (rowData.action === "syncEventToSupabase") {
      var ss2 = SpreadsheetApp.openById(SHEET_ID);
      var syncRes = syncEventToSupabase(ss2, rowData.eventName);
      return ContentService.createTextOutput(JSON.stringify(syncRes)).setMimeType(ContentService.MimeType.JSON);
    }
    if (rowData.action === "syncPurchaseRecords") {
      var ss3 = SpreadsheetApp.openById(SHEET_ID);
      syncPurchaseRecordsToSupabase(rowData.eventName);
      return ContentService.createTextOutput(JSON.stringify({ result: "ok" })).setMimeType(ContentService.MimeType.JSON);
    }
    if (rowData.action === "sendHKPickedUp") {
      var ss_ = SpreadsheetApp.openById(SHEET_ID);
      var puResult = sendHKPickedUpEmail(ss_, rowData.orderId);
      return ContentService.createTextOutput(JSON.stringify(puResult)).setMimeType(ContentService.MimeType.JSON);
    }
    if (rowData.action === "sendHKShipping") {
      var shResult = sendHKShippingEmail(ss, rowData.orderId, rowData.trackingNo || "");
      return ContentService.createTextOutput(JSON.stringify(shResult)).setMimeType(ContentService.MimeType.JSON);
    }
    if (rowData.action === "sendHKDelivered") {
      var sdResult = sendHKDeliveredEmail(ss, rowData.orderId, rowData.trackingNo || "");
      return ContentService.createTextOutput(JSON.stringify(sdResult)).setMimeType(ContentService.MimeType.JSON);
    }
    if (rowData.action === "sendStockoutEmail") {
      var result = sendStockoutEmail(ss, rowData.orderId, rowData.outOfStockItems || []);
      return ContentService.createTextOutput(JSON.stringify(result))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 8：更新運單號碼 / 賣貨便網址 ──────────────────────────
    if (rowData.action === "updateShipmentRef") {
      var orderSheet = ss.getSheetByName("訂單紀錄");
      if (!orderSheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單紀錄" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var oRows = orderSheet.getDataRange().getValues();
      for (var r = 1; r < oRows.length; r++) {
        if (oRows[r][0].toString() === rowData.orderId.toString()) {
          orderSheet.getRange(r + 1, 19).setValue(rowData.shipmentRef || ""); // S欄 = 第19欄
          // 即時檢查是否4欄齊全 → 改狀態為「已到貨」
          var rData24 = orderSheet.getRange(r + 1, 1, 1, 24).getValues()[0];
          checkAndSetArrived(orderSheet, r, rData24);
          return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
                               .setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到訂單編號" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 7：儲存訂貨區已購買數量 ──────────────────────────────
    if (rowData.action === "savePurchased") {
      var evName = rowData.eventName ? rowData.eventName.toString().trim() : "";
      var purchItems = rowData.items;
      if (!evName || !Array.isArray(purchItems)) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "缺少 eventName 或 items" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      savePurchased(ss, evName, purchItems);
      return ContentService.createTextOutput(JSON.stringify({ result: "success", saved: purchItems.length }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 6：批量更新商品款式庫存（J欄文字格式）────────────────────
    if (rowData.action === "batchUpdateStockLimit") {
      var targetSh = rowData.sheetName ? rowData.sheetName.toString() : "";
      var updates = rowData.updates; // [{ row, subStockLimit: { 款式: 庫存 } 或 null }]
      if (!targetSh || !Array.isArray(updates)) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "缺少 sheetName 或 updates" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      var bSheet = ss.getSheetByName(targetSh);
      if (!bSheet) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到分頁：" + targetSh }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      for (var ui = 0; ui < updates.length; ui++) {
        var upd = updates[ui];
        var rowNum = parseInt(upd.row);
        // subStockLimit 是物件，序列化成 "白色:1,黑色:0" 格式
        var val = serializeSubStockLimit(upd.subStockLimit);
        bSheet.getRange(rowNum, 10).setValue(val); // J欄 = 第10欄
      }
      return ContentService.createTextOutput(JSON.stringify({ result: "success", updated: updates.length }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── Nissen 代購訂單確認 email（admin 從 admin panel 觸發）──
    if (rowData.action === "sendNissenConfirmation") {
      try {
        var cnOrder = rowData.order || {};
        if (!cnOrder.customer_email || !cnOrder.order_no) {
          return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "缺少 order 資料" }))
                               .setMimeType(ContentService.MimeType.JSON);
        }
        var cnSubject = "[852hk.81jp] 訂單確認 " + cnOrder.order_no;
        var cnBody = buildNissenEmailHtml_(cnOrder, true);
        MailApp.sendEmail({ to: cnOrder.customer_email, subject: cnSubject, htmlBody: cnBody });
        return ContentService.createTextOutput(JSON.stringify({ result: "ok" }))
                             .setMimeType(ContentService.MimeType.JSON);
      } catch(cnErr) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: cnErr.toString() }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── Nissen 訂單通知（前台下單後 POST fire-and-forget，通知 admin）──
    if (rowData.action === "nissenOrderNotify") {
      try {
        var pnOrder = rowData.order || {};
        if (pnOrder.order_no || pnOrder.orderNo) {
          var pnSubject = "[新 Nissen 代購訂單] " + (pnOrder.order_no || pnOrder.orderNo) + " — " + (pnOrder.customer_name || pnOrder.customerName || "") + " (" + (pnOrder.region || "").toUpperCase() + ")";
          MailApp.sendEmail({ to: MY_NOTIFICATION_EMAIL, subject: pnSubject, htmlBody: buildNissenEmailHtml_(pnOrder, false) });
        }
      } catch(pnErr) { Logger.log("nissenOrderNotify(POST) error: " + pnErr); }
      return ContentService.createTextOutput(JSON.stringify({ result: "ok" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 現貨搶購付款確認 email（admin 從 admin panel 觸發）──
    if (rowData.action === "sendFlashConfirmation") {
      try {
        var fcOrder = rowData.order || {};
        if (!fcOrder.email) {
          return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "訂單沒有 email" }))
                               .setMimeType(ContentService.MimeType.JSON);
        }
        var fcRef = (fcOrder.token || "").slice(0, 8).toUpperCase();
        var fcIsTW = (fcOrder.region === 'tw');
        var fcSubject = "[852hk.81jp] ⚡ 現貨訂單付款確認 " + fcRef;
        var fcBody =
          '<div style="font-family:sans-serif;max-width:520px;margin:0 auto;">' +
          '<h2 style="color:#10b981;">✅ 已確認收到您的款項！</h2>' +
          '<p style="font-size:14px;line-height:1.8;">' + (fcOrder.customer_name || "") + ' 您好，</p>' +
          '<p style="font-size:14px;line-height:1.8;">我們已確認收到您的付款，以下現貨已為您保留：</p>' +
          '<div style="background:#f7f6f3;border-radius:10px;padding:14px 18px;margin:14px 0;">' +
          '<p style="font-size:15px;font-weight:700;margin:0;">' + (rowData.item_name || "現貨商品") + '</p>' +
          '<p style="font-size:12px;color:#888;margin:6px 0 0;">訂單參考編號：<strong>' + fcRef + '</strong></p>' +
          '</div>' +
          '<p style="font-size:14px;line-height:1.8;">' +
          (fcIsTW ? '我們會盡快安排出貨（賣貨便），出貨後會再通知您。' : '我們會盡快與您確認取貨／寄送安排。') +
          '</p>' +
          '<p style="font-size:13px;line-height:1.8;">查詢訂單進度：<a href="https://jpshopper-852hk81jp.vercel.app/status?id=' + fcRef + '" style="color:#2563eb;">按此查詢</a></p>' +
          '<p style="font-size:12px;color:#888;margin-top:20px;">如有任何問題，歡迎直接回覆此 Email 或 IG 私訊我們。<br>852hk.81jp</p>' +
          '</div>';
        MailApp.sendEmail({ to: fcOrder.email, subject: fcSubject, htmlBody: fcBody });
        return ContentService.createTextOutput(JSON.stringify({ result: "ok" }))
                             .setMimeType(ContentService.MimeType.JSON);
      } catch(fcErr) {
        return ContentService.createTextOutput(JSON.stringify({ result: "error", message: fcErr.toString() }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── 出 post 草稿存入「出Post」分頁（俾 Make 讀去自動出文）──
    if (rowData.action === "savePostDraft") {
      try {
        // 圖片除咗夾埋一格（F），仲逐張開一欄（L 之後）——
        // Make 要砌 carousel 嘅話，逐欄 map 落每一格，
        // 好過喺佢度將一格拆成 array，少一步易錯嘅嘢。
        var PD_MAX_IMG = 10;
        var pdHeader = [
          "建立時間", "地區", "店舖", "件數", "商品名", "圖片網址（全部）",
          "商品網址", "文案", "Hashtag", "狀態", "出咗嘅時間"
        ];
        for (var hi = 1; hi <= PD_MAX_IMG; hi++) pdHeader.push("圖片" + hi);

        var pdSheet = ss.getSheetByName("出Post");
        if (!pdSheet) {
          pdSheet = ss.insertSheet("出Post");
          pdSheet.appendRow(pdHeader);
          pdSheet.setFrozenRows(1);
          pdSheet.getRange(1, 1, 1, pdHeader.length).setFontWeight("bold");
        } else if (pdSheet.getLastColumn() < pdHeader.length) {
          // 舊版得 11 欄，補返啲圖片欄個標題
          pdSheet.getRange(1, 1, 1, pdHeader.length).setValues([pdHeader])
                 .setFontWeight("bold");
        }

        // Make 讀一行就夠砌一個 post
        var pdItems  = rowData.items || [];
        var pdPics   = pdItems.map(function(i){ return i.image || ""; }).filter(String);
        var pdUrls   = pdItems.map(function(i){ return i.url || ""; })
                              .filter(String).join("\n");
        var pdNames  = pdItems.map(function(i){ return i.name || ""; })
                              .filter(String).join("\n");

        var pdRow = [
          new Date(),
          (rowData.region === 'tw') ? '🇹🇼 台灣' : '🇭🇰 香港',
          rowData.shops   || "",
          pdItems.length,
          pdNames,
          pdPics.join("\n"),
          pdUrls,
          rowData.caption || "",
          rowData.tags    || "",
          "待出",
          ""
        ];
        for (var pi2 = 0; pi2 < PD_MAX_IMG; pi2++) pdRow.push(pdPics[pi2] || "");

        pdSheet.appendRow(pdRow);

        return ContentService.createTextOutput(JSON.stringify({
          result: "ok", row: pdSheet.getLastRow(), count: pdItems.length
        })).setMimeType(ContentService.MimeType.JSON);
      } catch(pdErr) {
        return ContentService.createTextOutput(JSON.stringify({
          result: "error", message: pdErr.toString()
        })).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ── 現貨搶購訂單通知（flash.html 提交表單後觸發，通知 admin）──
    if (rowData.action === "flashOrderNotify") {
      try {
        var flRegionTxt = (rowData.region === 'tw') ? '🇹🇼 台灣' : '🇭🇰 香港';
        var flSubject = "[⚡現貨搶購] " + (rowData.item_name || "") + " — " + (rowData.customer_name || "");
        var flBody =
          '<div style="font-family:sans-serif;max-width:520px;">' +
          '<h2 style="color:#c2410c;">⚡ 新現貨搶購訂單</h2>' +
          '<table style="border-collapse:collapse;font-size:14px;">' +
          '<tr><td style="padding:4px 12px 4px 0;color:#888;">商品</td><td><strong>' + (rowData.item_name || "") + '</strong></td></tr>' +
          '<tr><td style="padding:4px 12px 4px 0;color:#888;">地區</td><td>' + flRegionTxt + '</td></tr>' +
          '<tr><td style="padding:4px 12px 4px 0;color:#888;">姓名</td><td>' + (rowData.customer_name || "") + '</td></tr>' +
          '<tr><td style="padding:4px 12px 4px 0;color:#888;">電話</td><td>' + (rowData.phone || "") + '</td></tr>' +
          '<tr><td style="padding:4px 12px 4px 0;color:#888;">Email</td><td>' + (rowData.email || "") + '</td></tr>' +
          '<tr><td style="padding:4px 12px 4px 0;color:#888;">付款方式</td><td>' + (rowData.pay_method || "") + '</td></tr>' +
          (rowData.remark ? '<tr><td style="padding:4px 12px 4px 0;color:#888;">備註</td><td>' + rowData.remark + '</td></tr>' : '') +
          '<tr><td style="padding:4px 12px 4px 0;color:#888;">Token</td><td style="font-size:12px;color:#aaa;">' + (rowData.token || "") + '</td></tr>' +
          '</table>' +
          '<p style="margin-top:16px;"><a href="' + ADMIN_PAGE_URL + '" style="color:#2563eb;">前往 Admin 後台處理 →</a></p>' +
          '</div>';
        MailApp.sendEmail({ to: MY_NOTIFICATION_EMAIL, subject: flSubject, htmlBody: flBody });
      } catch(flErr) { Logger.log("flashOrderNotify error: " + flErr); }
      return ContentService.createTextOutput(JSON.stringify({ result: "ok" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ── 功能 4：消費者下單 ────────────────────────────────────────────
    var sheet = ss.getSheetByName("訂單紀錄");
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ result: "error", message: "找不到『訂單紀錄』分頁" })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var timestamp = new Date();
    var orderId = "ORD" + timestamp.getTime();
    var finalEventName = rowData.eventSheetName || rowData.targetSheet || "現場快閃代購";
    var orderRegion = rowData.region ? rowData.region.toString().toLowerCase() : "hk";

    // ── 計算付款期限提示文字 ──
    // 規則：訂單時間 + 20分鐘 vs 活動截止時間，取較早者
    var deadlineMap = syncAndGetDeadlines(ss);
    var eventDeadlineStr = deadlineMap[finalEventName] || "";
    var twentyMinLater = new Date(timestamp.getTime() + 20 * 60 * 1000);
    var payDeadlineText = "";
    var payDeadlineTextTW = "";

    var payDueDate = twentyMinLater; // 預設用 +20分鐘
    if (eventDeadlineStr) {
      var eventDeadline = new Date(eventDeadlineStr.replace(/\//g, "-"));
      if (!isNaN(eventDeadline) && eventDeadline < twentyMinLater) {
        payDueDate = eventDeadline; // 活動截止更早，用活動截止
      }
    }
    var pdMonth = payDueDate.getMonth() + 1;
    var pdDate  = payDueDate.getDate();
    var pdHour  = payDueDate.getHours();
    var pdMin   = payDueDate.getMinutes() < 10 ? "0" + payDueDate.getMinutes() : payDueDate.getMinutes();
    var pdStr   = pdMonth + "月" + pdDate + "日 " + pdHour + ":" + pdMin;
    payDeadlineText   = "請於" + pdStr + " 前完成付款";
    payDeadlineTextTW = "請於" + pdStr + " 前完成匯款";
    
    var newRow = [
      orderId, timestamp, finalEventName, rowData.itemsSummary || "",
      rowData.totalAmount || 0, rowData.shippingFee || 0, rowData.totalJPYAmount || 0,
      rowData.paymentMethod || "", rowData.recipientName || "", rowData.recipientPhone || "",
      rowData.recipientEmail || "", rowData.pickupType || "", rowData.pickupCode || "",
      rowData.pickupName || "", rowData.pickupAddress || "", "待處理"
    ];
    sheet.appendRow(newRow);

    // 同步訂單數量到 Supabase（即時庫存更新）
    try {
      var sbEventName = rowData.eventName || "";
      var sbSummary   = rowData.itemsSummary || rowData.cartSummary || "";
      if (sbEventName && sbSummary) updateOrderedQtyInSupabase(sbEventName, sbSummary, false);
    } catch(sbErr) { Logger.log("Supabase qty sync err: " + sbErr); }
    
    // ── 建立商品圖片 HTML（HK版 email 使用）────────────────────────────
    var eventSheet = getSheetByEventName(ss, finalEventName);
    var backupProducts = [];
    if (eventSheet) {
      backupProducts = eventSheet.getDataRange().getValues();
    }

    var finalItemsList = [];
    if (rowData.cartItems && Array.isArray(rowData.cartItems) && rowData.cartItems.length > 0) {
      finalItemsList = rowData.cartItems;
    } else {
      var summaryText = rowData.itemsSummary ? rowData.itemsSummary.toString() : "";
      var regex = /([^,，(（]+)(?:[(（]([^)）]+)[)）])?/g;
      var match;
      while ((match = regex.exec(summaryText)) !== null) {
        var pName = match[1].replace(/^[,\s\n\r]+|[,\s\n\r]+$/g, "").trim();
        var insideBrackets = match[2] ? match[2].trim() : "";
        if (!pName) continue;
        if (insideBrackets) {
          var subParts = insideBrackets.split(/[,，]/);
          for (var s = 0; s < subParts.length; s++) {
            var part = subParts[s].trim();
            var pSub = part;
            var pQty = 1;
            if (part.indexOf("x") !== -1) {
              pSub = part.split("x")[0].trim();
              pQty = parseInt(part.split("x")[1]) || 1;
            }
            finalItemsList.push({ name: pName, sub: pSub, qty: pQty, price: null, photo: null });
          }
        } else {
          finalItemsList.push({ name: pName, sub: "", qty: 1, price: null, photo: null });
        }
      }
    }

    // 建立 HK 版商品表格 HTML（用 HKD 價格）
    var tableRowsHtml = "";
    for (var k = 0; k < finalItemsList.length; k++) {
      var item = finalItemsList[k];
      var displayName = item.name + (item.sub ? " (" + item.sub + ")" : "");
      var quantity = parseInt(item.qty) || 1;
      var imgUrl = item.photo || "";
      var currentPrice = parseFloat(item.price) || 0;
      
      if (!imgUrl || currentPrice === 0) {
        var cleanItemName = item.name.toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5]/g, "");
        for (var p = 1; p < backupProducts.length; p++) {
          var sheetProdName = backupProducts[p][0] ? backupProducts[p][0].toString().trim() : "";
          var sheetProdSub = backupProducts[p][1] ? backupProducts[p][1].toString().trim() : "";
          var cleanSheetName = sheetProdName.toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5]/g, "");
          if (cleanItemName.indexOf(cleanSheetName) !== -1 || cleanSheetName.indexOf(cleanItemName) !== -1) {
            if (!imgUrl && backupProducts[p][5]) imgUrl = backupProducts[p][5].toString().trim();
            if (currentPrice === 0 && backupProducts[p][4]) currentPrice = parseFloat(backupProducts[p][4]);
            if (item.sub && sheetProdSub) {
              var cleanItemSub = item.sub.toLowerCase().replace(/[^a-z0-9]/g, "");
              var cleanSheetSub = sheetProdSub.toLowerCase().replace(/[^a-z0-9]/g, "");
              if (cleanItemSub.indexOf(cleanSheetSub) !== -1 || cleanSheetSub.indexOf(cleanItemSub) !== -1) {
                if (backupProducts[p][5]) imgUrl = backupProducts[p][5].toString().trim();
                if (backupProducts[p][4]) currentPrice = parseFloat(backupProducts[p][4]);
                break;
              }
            }
          }
        }
      }
      var hasValidImg = imgUrl && imgUrl.indexOf("http") === 0;

      // 根據地區顯示對應幣別金額
      var displayCurrency, itemSubtotal;
      if (orderRegion === "tw") {
        // TW：從 cartItems 讀 twd 價格；若找不到則從 Sheet I欄補查
        var twdPrice = parseFloat(item.twd) || 0;
        if (twdPrice === 0 && backupProducts.length > 0) {
          var cleanIN = item.name ? item.name.toString().toLowerCase().replace(/[^a-z0-9一-龥]/g, "") : "";
          for (var bp = 1; bp < backupProducts.length; bp++) {
            var bpN = backupProducts[bp][0] ? backupProducts[bp][0].toString().trim() : "";
            var bpClean = bpN.toLowerCase().replace(/[^a-z0-9一-龥]/g, "");
            if (cleanIN && bpClean && (cleanIN.indexOf(bpClean) !== -1 || bpClean.indexOf(cleanIN) !== -1)) {
              twdPrice = parseFloat(backupProducts[bp][8]) || 0; // I欄
              break;
            }
          }
        }
        itemSubtotal = twdPrice * quantity;
        displayCurrency = "NT$";
      } else {
        itemSubtotal = currentPrice * quantity;
        displayCurrency = "HK$";
      }

      tableRowsHtml += `
        <tr style="border-bottom: 1px solid #eeeeee;">
          <td style="padding: 12px 10px; font-size: 14px; color: #333333; padding-left: 15px;">
            <table style="width: 100%; border-collapse: collapse; border: none;"><tr>
              ${hasValidImg ? `<td style="padding: 0; width: 50px; vertical-align: middle;"><img src="${imgUrl}" style="width: 50px; height: 50px; object-fit: cover; border-radius: 4px; border: 1px solid #eef0ed; display: block;" /></td>` : ""}
              <td style="padding: 0; padding-left: ${hasValidImg ? "12px" : "0"}; vertical-align: middle; line-height: 1.4;">${displayName}</td>
            </tr></table>
          </td>
          <td style="padding: 12px 10px; font-size: 14px; color: #555555; text-align: center;">${quantity}</td>
          <td style="padding: 12px 10px; font-size: 14px; color: #111111; text-align: right; font-weight: 500; padding-right: 15px; width: 80px;">
            ${displayCurrency} ${itemSubtotal > 0 ? itemSubtotal : "—"}
          </td>
        </tr>`;
    }

    // ── 📢 管理員通知信（標明 HK / TW）───────────────────────────────
    try {
      var adminEmailBody = `
        <div style="font-family: sans-serif; padding: 20px; border: 1px solid #ddd; border-radius: 6px; max-width: 600px; background-color: #ffffff;">
          <p style="font-size: 16px; font-weight: bold; color: #111;">📢 您有新訂單進來囉！</p>
          <p><strong>所屬活動：</strong> ${finalEventName}</p>
          <p><strong>地區版本：</strong> ${orderRegion.toUpperCase()}</p>
          <p><strong>訂單號碼：</strong> ${orderId}</p>
          <p><strong>付款方式：</strong> ${rowData.paymentMethod || ""}</p>
          <hr style="border: none; border-top: 1px solid #eee; margin: 15px 0;">
          <h4 style="margin-bottom: 8px; color: #555;">🛒 訂單商品明細：</h4>
          <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
            <tr style="background: #f5f5f5; border-bottom: 1px solid #ddd;">
              <th style="padding:8px; text-align:left;">商品</th>
              <th style="padding:8px; text-align:center; width:50px;">數量</th>
            </tr>
            ${tableRowsHtml}
          </table>
          <p style="text-align: right; margin-top: 10px;"><strong>郵費：</strong> ${orderRegion === "tw" ? "NT$ 待通知" : "HK$ " + (rowData.shippingFee || 0)}</p>
          <p style="text-align: right; font-size: 15px; color: #ff4d4d;"><strong>訂單總額：</strong> ${orderRegion === "tw" ? "NT$ " : "HK$ "}${rowData.totalAmount}</p>
          <hr style="border: none; border-top: 1px solid #eee; margin: 15px 0;">
          <p><strong>收件人：</strong> ${rowData.recipientName} (${rowData.recipientPhone})</p>
          <p><strong>取件方式：</strong> ${rowData.pickupName || "賣貨便（待通知）"} ${rowData.pickupAddress ? "- " + rowData.pickupAddress : ""}</p>
          <br>
          <a href="${ADMIN_PAGE_URL}" target="_blank" style="display: inline-block; background-color: #111; color: #fff; padding: 10px 15px; text-decoration: none; border-radius: 4px; font-size: 14px;">進入後台核對款項</a>
        </div>`;

      MailApp.sendEmail({
        to: MY_NOTIFICATION_EMAIL,
        subject: `📢 [${finalEventName}][${orderRegion.toUpperCase()}] 收到新訂單！編號：${orderId}`,
        htmlBody: adminEmailBody
      });
    } catch(emailErr) {
      Logger.log("管理員信件發送失敗: " + emailErr.toString());
    }
    
    // ── ✉️ 消費者通知信（HK / TW 分流）─────────────────────────────
    var targetEmail = rowData.recipientEmail ? rowData.recipientEmail.toString().trim() : "";
    
    if (targetEmail !== "") {

      // ════════════════════════════════════════════════════════════════
      // 🇹🇼 台灣版 email
      // ════════════════════════════════════════════════════════════════
      if (orderRegion === "tw") {
        var twOrderId = orderId;
        var twEventName = finalEventName;
        var twTotal = rowData.totalAmount || 0;
        // 手機末3碼
        var twPhone = rowData.recipientPhone ? rowData.recipientPhone.toString() : "";
        var twPhoneLast3 = twPhone.length >= 3 ? twPhone.slice(-3) : twPhone;
        var twPayMethod  = rowData.paymentMethod ? rowData.paymentMethod.toString() : "";
        var isLinePay    = twPayMethod.toUpperCase().indexOf("LINE") !== -1;

        // 🖼️ 從 Sheet F欄預建圖片對照表（商品名 → 圖片網址）
        var twPhotoMap = {};
        for (var bp = 1; bp < backupProducts.length; bp++) {
          var bpName = backupProducts[bp][0] ? backupProducts[bp][0].toString().trim() : "";
          var bpPhoto = backupProducts[bp][5] ? backupProducts[bp][5].toString().trim() : "";
          if (bpName && bpPhoto) twPhotoMap[bpName] = bpPhoto;
        }

        // 台灣版商品明細（顯示 TWD，圖片從 Sheet F欄查）
        var twTableRowsHtml = "";
        var twItemsList = (rowData.cartItems && Array.isArray(rowData.cartItems)) ? rowData.cartItems : [];
        for (var ti = 0; ti < twItemsList.length; ti++) {
          var tItem = twItemsList[ti];
          var tDisplayName = tItem.name + (tItem.sub ? " (" + tItem.sub + ")" : "");
          var tQty = parseInt(tItem.qty) || 1;
          var tPrice = parseFloat(tItem.twd) || 0;
          var tSubtotal = tPrice * tQty;

          // 優先用 cartItem 傳來的 photo，再從 Sheet F欄查，找不到就不顯示圖片
          var tItemName = tItem.name ? tItem.name.toString().trim() : "";
          var tImgUrl = (tItem.photo && tItem.photo.indexOf("http") === 0) ? tItem.photo : (twPhotoMap[tItemName] || "");
          if (!tImgUrl || tImgUrl.indexOf("http") !== 0) {
            var tCleanName = tItemName.toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5]/g, "");
            for (var bpKey in twPhotoMap) {
              var bpClean = bpKey.toLowerCase().replace(/[^a-z0-9\u4e00-\u9fa5]/g, "");
              if (tCleanName && bpClean && (tCleanName.indexOf(bpClean) !== -1 || bpClean.indexOf(tCleanName) !== -1)) {
                tImgUrl = twPhotoMap[bpKey];
                break;
              }
            }
          }
          var tHasValidImg = tImgUrl && tImgUrl.indexOf("http") === 0;
          twTableRowsHtml += `
            <tr style="border-bottom: 1px solid #eeeeee;">
              <td style="padding: 12px 10px 12px 15px; font-size: 14px; color: #333333;">
                <table style="width:100%; border-collapse:collapse; border:none;"><tr>
                  ${tHasValidImg ? `<td style="padding:0; width:50px; vertical-align:middle;"><img src="${tImgUrl}" style="width:50px; height:50px; object-fit:cover; border-radius:4px; border:1px solid #eef0ed; display:block;" /></td>` : ""}
                  <td style="padding:0 ${tHasValidImg ? "0 0 12px" : ""}; vertical-align:middle; line-height:1.4;">${tDisplayName}</td>
                </tr></table>
              </td>
              <td style="padding:12px 10px; font-size:14px; color:#555555; text-align:center;">${tQty}</td>
              <td style="padding:12px 15px 12px 10px; font-size:14px; color:#2980b9; text-align:right; font-weight:500;">
                ${tSubtotal > 0 ? "NT$ " + tSubtotal : "—"}
              </td>
            </tr>`;
        }

        var twEmailSubject = `[已收到訂單] 代購 ${twEventName} — 886tw.81jp 台灣`;
        var twEmailBody = `
          <div style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width:600px; margin:0 auto; color:#333333; line-height:1.6; padding:20px; background-color:#f7f6f3;">
            <h2 style="font-size:18px; font-weight:500; border-bottom:1px solid #e5e5e5; padding-bottom:10px; color:#1a1a1a;">
              代購 ${twEventName}
            </h2>
            <p style="font-size:14px;">您好！感謝您的訂購 🎉</p><p style="font-size:14px; margin-top:12px; color:#555;">⚠️ 我們會在收取款項後為客人保留商品✨</p><p style="font-size:13px; color:#888; margin-top:8px; line-height:1.7;">若未於20分鐘內付款，將關閉訂單，您可以在活動結束前再次下單。感謝您的確認🙇🏻‍♀️</p><p style="font-size:13px; color:#888; margin-top:6px; line-height:1.7;">如遇缺貨，將以付款順序分配商品。</p>
            <p style="font-size:14px; background:#f2f2f2; padding:8px 12px; border-radius:4px; display:inline-block;">
              <strong>訂單編號：</strong> ${twOrderId}
            </p>

            <!-- 進度條 -->
            <div style="margin:35px 0; padding:0 10px;">
              <div style="position:relative; display:flex; justify-content:space-between; align-items:center; width:100%;">
                <div style="position:absolute; top:5px; left:0; right:0; height:1px; background-color:#e5e5e5; z-index:1;"></div>
                <div style="position:relative; z-index:2; text-align:center; width:25%;">
                  <div style="width:11px; height:11px; background-color:#2980b9; border-radius:50%; margin:0 auto; border:2px solid #ffffff; box-shadow:0 0 0 1px #2980b9;"></div>
                  <div style="font-size:12px; font-weight:bold; color:#2980b9; margin-top:8px;">已收到訂單</div>
                </div>
                <div style="position:relative; z-index:2; text-align:center; width:25%;">
                  <div style="width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;"></div>
                  <div style="font-size:11px; color:#888888; margin-top:8px;">已付款</div>
                </div>
                <div style="position:relative; z-index:2; text-align:center; width:25%;">
                  <div style="width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;"></div>
                  <div style="font-size:11px; color:#888888; margin-top:8px;">已到台</div>
                </div>
                <div style="position:relative; z-index:2; text-align:center; width:25%;">
                  <div style="width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;"></div>
                  <div style="font-size:11px; color:#888888; margin-top:8px;">已完成取貨</div>
                </div>
              </div>
            </div>

            <!-- 商品明細 -->
            <h3 style="font-size:14px; font-weight:bold; margin-top:25px; margin-bottom:10px; color:#444444;">📋 訂購商品明細</h3>
            <table style="width:100%; border-collapse:collapse; background:#ffffff; border:1px solid #eef0ed; border-radius:6px; overflow:hidden;">
              <thead>
                <tr style="background-color:#f0f7ff; border-bottom:1px solid #eef0ed;">
                  <th style="padding:10px 15px; text-align:left; font-size:13px; color:#666666; font-weight:500;">商品名稱與款式</th>
                  <th style="padding:10px; text-align:center; font-size:13px; color:#666666; font-weight:500; width:60px;">數量</th>
                  <th style="padding:10px 15px; text-align:right; font-size:13px; color:#666666; font-weight:500; width:90px;">小計</th>
                </tr>
              </thead>
              <tbody>${twTableRowsHtml}</tbody>
            </table>

            <p style="text-align:right; font-size:16px; font-weight:bold; margin-top:15px; color:#1a1a1a;">
              商品小計：<span style="font-size:20px; color:#2980b9;">NT$ ${twTotal}</span>
            </p>
            <p style="text-align:right; font-size:12px; color:#888; margin-top:4px;">
              ＋ 國際運費（NT$20／50g）＋ NT$38 賣貨便運費（出貨時另行通知）
            </p>

            <!-- 付款資訊 -->
            ${isLinePay ? `
            <div style="background-color:#f0fff4; border:1px solid #bbf7d0; border-radius:6px; padding:14px; margin:20px 0; font-size:12px; color:#166534; line-height:1.9;">
              <div style="font-weight:700; font-size:15px; margin-bottom:8px;">💚 LINE PAY 付款步驟</div>
              <div>① 完成訂單後，點擊連結加入我們的 LINE</div>
              <div>② 完成付款後，在聊天頁面複製以下格式並傳送給我們：</div>
              <div style="background:#fff; border:1px solid #86efac; border-radius:6px; padding:12px 14px; font-family:monospace; font-size:13px; line-height:2.4; color:#111; white-space:pre-line; margin:10px 0;">訂單編號：${twOrderId}
手機末3碼：${twPhoneLast3}
金額：NT$ ${twTotal}
已付款💰</div>
              <a href="${LINE_URL}" target="_blank" style="display:inline-block; background:#00B900; color:#fff; font-size:13px; font-weight:700; padding:9px 22px; border-radius:20px; text-decoration:none; margin-top:4px;">💚 加入 LINE 並付款</a>
              <p style="margin:10px 0 0; font-size:12px; color:#166534; font-weight:700;">⏰ ${payDeadlineTextTW}，逾時訂單將自動取消。</p>
            </div>` : `
            <div style="background-color:#f0f7ff; border-left:4px solid #2980b9; padding:15px; margin:20px 0; border-radius:4px;">
              <p style="margin:0; font-weight:bold; color:#1565c0; font-size:15px;">💰 郵局銀行匯款資訊</p>
              <p style="margin:6px 0 10px; font-size:15px; color:#1a237e;">匯款金額：<strong style="font-size:20px;">NT$ ${twTotal}</strong></p>
              <p style="margin:10px 0 4px; font-size:14px; color:#333;">銀行：<strong>中華郵政（郵局）</strong></p>
              <p style="margin:4px 0; font-size:14px; color:#333;">局號／帳號：<strong>0041860-0025565</strong></p>
              <p style="margin:4px 0; font-size:14px; color:#333;">戶名：<strong>周◯恩</strong></p>
              <p style="margin:12px 0 6px; font-size:14px; font-weight:700; color:#1565c0;">
                ⚠️ ${payDeadlineTextTW}，匯款時請於備註欄填寫您的<strong>手機末3碼</strong>：
              </p>
              <span style="display:inline-block; background:#fff; border:2px dashed #2980b9; padding:6px 18px; border-radius:6px; font-family:monospace; font-size:22px; font-weight:bold; color:#2980b9; margin:4px 0 10px; letter-spacing:6px;">${twPhoneLast3}</span>
              <p style="margin:8px 0 0; font-size:12px; color:#888; line-height:1.7;">
                匯款後如24小時內沒有收到已付款通知，請主動回覆電郵或<a href="https://www.threads.com/@886tw.81jp" style="color:#2980b9; font-weight:600;">私訊我們</a>。
              </p>
            </div>`}

            <!-- 取貨流程 -->
            <h3 style="font-size:14px; font-weight:bold; margin-top:25px; margin-bottom:10px; color:#444444;">📦 取貨流程說明</h3>
            <div style="background:#ffffff; border:1px solid #eef0ed; padding:16px; border-radius:6px; font-size:13px; color:#555555; line-height:2.1;">
              <div>① 商品自日本購入後集貨寄送至台灣</div>
              <div>② 商品到達台灣並稱重後，店主將以 Email 通知您，並於賣貨便建立出貨單</div>
              <div>③ 收到通知後，請於賣貨便選擇超商門市，並完成<strong>國際運費 + NT$38</strong> 的付款</div>
              <div>④ 到貨後前往超商取貨，取貨付款即完成！</div>
            </div>
            <div style="margin-top:10px; padding:10px 14px; background:#e3f2fd; border-radius:6px; font-size:12px; color:#555; line-height:1.8;">
              💡 <strong>運費計算：</strong>NT$20 / 50g，不足 50g 以 50g 計算，超過 50g 以實重計算。具體金額將於出貨 Email 告知。
            </div>

            <div style="margin-top:30px; padding-top:15px; border-top:1px dashed #dddddd; font-size:12px; color:#777777; line-height:1.8;">
              <p>✅ 我們收到匯款後才會進行採購，並發送付款確認 Email 通知您。</p>
              <p>如有任何問題，歡迎直接回覆此 Email 或於 Threads 私訊我們查詢，謝謝！</p>
            </div>
            <div style="margin-top:30px; font-size:13px; color:#333333; line-height:1.5;">
              <strong>886tw.81jp</strong><br>
              <a href="https://www.threads.com/@886tw.81jp?igshid=NTc4MTIwNjQ2YQ==" style="color:#2980b9;">Threads @886tw.81jp</a>　<a href="https://www.instagram.com/886tw.81jp?igsh=MW8zMmVncGVwNmd1dg%3D%3D&utm_source=qr" style="color:#e1306c;">Instagram @886tw.81jp</a>
            </div>
          </div>`;

        try {
          MailApp.sendEmail({
            to: targetEmail,
            subject: twEmailSubject,
            body: "感謝您的訂購！訂單編號：" + twOrderId + "。請查看 HTML 郵件以獲取付款與取貨資訊。",
            htmlBody: twEmailBody
          });
        } catch(twMailErr) {
          Logger.log("TW 消費者信件發送失敗: " + twMailErr.toString());
        }

      // ════════════════════════════════════════════════════════════════
      // 🇭🇰 香港版 email（原有邏輯完整保留）
      // ════════════════════════════════════════════════════════════════
      } else {
        var copyCodeHtml = `<span style="display: inline-block; background-color: #ffffff; border: 1px dashed #cccccc; padding: 4px 10px; border-radius: 4px; font-family: monospace; font-size: 15px; color: #111111; font-weight: bold; margin: 2px 5px;">${orderId}</span>`;
        var paymentDetailsHtml = "";
        var payMethod = rowData.paymentMethod ? rowData.paymentMethod.toString().toUpperCase() : "";
        
        if (payMethod.indexOf("PAYME") !== -1) {
          paymentDetailsHtml = `
            <div style="background-color: #fff5f5; border-left: 4px solid #ff4d4d; padding: 15px; margin: 15px 0; border-radius: 4px;">
              <p style="margin: 0; font-weight: bold; color: #d93838; font-size: 15px;">💰 PayMe 付款資料：</p>
              <p style="margin: 6px 0 10px; font-size: 15px; color: #b71c1c;">匯款金額：<strong style="font-size: 20px;">HK$ ${rowData.totalAmount}</strong></p>
              <p style="margin: 8px 0;">
                <a href="https://payme.hsbc/miru" target="_blank" style="display:inline-block; background-color:#ff4d4d; color:#ffffff; font-weight:bold; font-size:15px; padding:10px 20px; border-radius:6px; text-decoration:none;">💳 一按即 PayMe</a>
              </p>
              <p style="margin: 6px 0 0; font-size: 12px; color: #999;">或複製連結：https://payme.hsbc/miru</p>
              <p style="margin: 8px 0 0 0; font-size: 13px; color: #666666;">⚠️ ${payDeadlineText}，付款時請於【備註】輸入訂單編號：${copyCodeHtml}</p>
            </div>`;
        } else if (payMethod.indexOf("FPS") !== -1 || payMethod.indexOf("轉數快") !== -1) {
          paymentDetailsHtml = `
            <div style="background-color: #f0f7ff; border-left: 4px solid #0066cc; padding: 15px; margin: 15px 0; border-radius: 4px;">
              <p style="margin: 0; font-weight: bold; color: #0052a3; font-size: 15px;">💰 轉數快 (FPS) 付款資料：</p>
              <p style="margin: 6px 0 10px; font-size: 15px; color: #1a237e;">匯款金額：<strong style="font-size: 20px;">HK$ ${rowData.totalAmount}</strong></p>
              <p style="margin: 8px 0; font-size: 14px; color: #333333;">轉數快快速支付號碼：<strong>8890873</strong><br>帳戶持有人：<strong>Chow W. Y.</strong></p>
              <p style="margin: 8px 0 0 0; font-size: 13px; color: #666666;">⚠️ ${payDeadlineText}，付款時請於【備註】輸入訂單編號：${copyCodeHtml}</p>
            </div>`;
        }
        
        var emailSubject = `[已收到訂單] ${finalEventName} - 852hk.81jp`;
        var emailBodyHtml = `
          <div style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width: 600px; margin: 0 auto; color: #333333; line-height: 1.6; padding: 20px; background-color: #fbfbfa;">
            <h2 style="font-size: 18px; font-weight: 500; border-bottom: 1px solid #e5e5e5; padding-bottom: 10px; color: #111111;">代購 ${finalEventName}</h2>
            <p style="font-size: 14px;">你好～ 感謝你落單 🎉</p><p style="font-size: 14px; margin-top:12px; color:#555;">⚠️ 我哋會喺收到款項後先幫客人保留商品✨</p><p style="font-size: 13px; color:#888; margin-top:8px; line-height:1.7;">若未能喺20分鐘內完成付款，我哋將關閉訂單。你可以於活動完結前重新落單，感謝你嘅確認🙇🏻‍♀️</p><p style="font-size:13px; color:#888; margin-top:6px; line-height:1.7;">如遇缺貨，將以付款先後次序分配商品。</p>
            <p style="font-size: 14px; background: #f2f2f2; padding: 8px 12px; border-radius: 4px; display: inline-block;"><strong>訂單編號：</strong> ${orderId}</p>
            
            <div style="margin: 35px 0; padding: 0 10px;">
              <div style="position: relative; display: flex; justify-content: space-between; align-items: center; width: 100%;">
                <div style="position: absolute; top: 5px; left: 0; right: 0; height: 1px; background-color: #e5e5e5; z-index: 1;"></div>
                <div style="position: relative; z-index: 2; text-align: center; width: 25%;">
                  <div style="width: 11px; height: 11px; background-color: #111111; border-radius: 50%; margin: 0 auto; border: 2px solid #ffffff; box-shadow: 0 0 0 1px #111111;"></div>
                  <div style="font-size: 12px; font-weight: bold; color: #111111; margin-top: 8px;">已收到訂單</div>
                </div>
                <div style="position: relative; z-index: 2; text-align: center; width: 25%;">
                  <div style="width: 11px; height: 11px; background-color: #d1d5db; border-radius: 50%; margin: 0 auto; border: 2px solid #ffffff;"></div>
                  <div style="font-size: 11px; color: #888888; margin-top: 8px;">已付款</div>
                </div>
                <div style="position: relative; z-index: 2; text-align: center; width: 25%;">
                  <div style="width: 11px; height: 11px; background-color: #d1d5db; border-radius: 50%; margin: 0 auto; border: 2px solid #ffffff;"></div>
                  <div style="font-size: 11px; color: #888888; margin-top: 8px;">到港途中</div>
                </div>
                <div style="position: relative; z-index: 2; text-align: center; width: 25%;">
                  <div style="width: 11px; height: 11px; background-color: #d1d5db; border-radius: 50%; margin: 0 auto; border: 2px solid #ffffff;"></div>
                  <div style="font-size: 11px; color: #888888; margin-top: 8px;">已寄出</div>
                </div>
              </div>
            </div>

            <h3 style="font-size: 14px; font-weight: bold; margin-top: 25px; margin-bottom: 10px; color: #444444;">📋 訂購商品明細</h3>
            <table style="width: 100%; border-collapse: collapse; background: #ffffff; border: 1px solid #eef0ed; border-radius: 6px; overflow: hidden;">
              <thead>
                <tr style="background-color: #fafbfa; border-bottom: 1px solid #eef0ed;">
                  <th style="padding: 10px; text-align: left; font-size: 13px; color: #666666; padding-left: 15px;">商品名稱與款式</th>
                  <th style="padding: 10px; text-align: center; font-size: 13px; color: #666666; width: 60px;">數量</th>
                  <th style="padding: 10px; text-align: right; font-size: 13px; color: #666666; width: 80px; padding-right: 15px;">小計</th>
                </tr>
              </thead>
              <tbody>${tableRowsHtml}</tbody>
            </table>

            <p style="text-align: right; font-size: 14px; margin-top: 15px; color: #666666;">運費：<span>HK$ ${rowData.shippingFee || 0}</span></p>
            <p style="text-align: right; font-size: 16px; font-weight: bold; margin-top: 5px; color: #111111;">總金額 (含郵費)：<span style="font-size: 20px; color: #ff4d4d;">HK$ ${rowData.totalAmount}</span></p>

            <h3 style="font-size: 14px; font-weight: bold; margin-top: 25px; margin-bottom: 5px; color: #444444;">📍 易寄取收件資料</h3>
            <div style="background: #ffffff; border: 1px solid #eef0ed; padding: 15px; border-radius: 6px; font-size: 13px; color: #555555; line-height: 1.8;">
              <strong>收件人姓名：</strong> ${rowData.recipientName || "-"}<br>
              <strong>聯絡電話：</strong> ${rowData.recipientPhone || "-"}<br>
              <strong>自提點名稱：</strong> ${rowData.pickupName || "-"}<br>
              <strong>自提點地址：</strong> ${rowData.pickupAddress || "-"}
            </div>

            <p style="font-size: 14px; margin-top: 25px; font-weight: bold;">💡 如資料無誤，請使用下方資訊完成付款：</p>
            ${paymentDetailsHtml}

            <div style="margin-top: 30px; padding-top: 15px; border-top: 1px dashed #dddddd; font-size: 12px; color: #777777;">
              <p>✨ 確認付款後，你會再收到一封訂單確認 Email。</p>
              <p>如有任何問題，歡迎直接回覆此 Email 或到 Threads inbox 我哋查詢，謝謝！</p>
            </div>
          </div>`;

        try {
          MailApp.sendEmail({
            to: targetEmail,
            subject: emailSubject,
            body: "感謝您的訂購！訂單編號: " + orderId,
            htmlBody: emailBodyHtml
          });
        } catch(hkMailErr) {
          Logger.log("HK 消費者信件發送失敗: " + hkMailErr.toString());
        }
      }
    }
    
    // ── 自動更新購貨紀錄 F欄（未付款 +qty）──
    try {
      var orderItems = parseSummaryItems(rowData.itemsSummary || "");
      updatePurchaseRecordQty(ss, finalEventName, orderItems, 6, 1); // F欄+
    } catch(e) { Logger.log("購貨紀錄F欄更新失敗: " + e.toString()); }

    return ContentService.createTextOutput(JSON.stringify({ result: "success", orderId: orderId })).setMimeType(ContentService.MimeType.JSON);
                         
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// =================================================================
// sendPaidEmailToCustomer：後台標記已收款後發信（HK / TW 分流）
// =================================================================
function sendPaidEmailToCustomer(ss, rowValues) {
  var orderId = rowValues[0];
  var eventName = rowValues[2];
  var itemsSummary = rowValues[3];
  var totalAmount = rowValues[4];
  var paymentMethod = rowValues[7] ? rowValues[7].toString() : "";
  var recipientName = rowValues[8];
  var recipientPhone = rowValues[9];
  var targetEmail = rowValues[10] ? rowValues[10].toString().trim() : "";
  var pickupType = rowValues[11];
  var pickupCode = rowValues[12];
  var pickupName = rowValues[13];
  var pickupAddress = rowValues[14];

  var currentEventName = eventName ? eventName.toString().trim() : "現場快閃代購";

  // 判斷地區：台灣訂單的 pickupType 為「賣貨便」，或付款方式包含「郵局」
  var isTW = (pickupType && pickupType.toString().indexOf("賣貨便") !== -1)
          || (paymentMethod && paymentMethod.indexOf("郵局") !== -1);

  // 從 Sheet 讀取商品資料（HKD price + TWD price + photo）
  var productMap = {}; // { name: { photo, price, twd } }
  try {
    var productSheet = getSheetByEventName(ss, currentEventName);
    if (productSheet) {
      var productRows = productSheet.getDataRange().getValues();
      for (var m = 1; m < productRows.length; m++) {
        var pName = productRows[m][0] ? productRows[m][0].toString().trim() : "";
        var pPrice = productRows[m][4] ? parseFloat(productRows[m][4]) : 0;
        var pPhoto = productRows[m][5] ? productRows[m][5].toString().trim() : "";
        var pTwd   = productRows[m][8] ? parseFloat(productRows[m][8]) : 0; // I欄 TWD
        if (pName) productMap[pName] = { photo: pPhoto, price: pPrice, twd: pTwd };
      }
    }
  } catch (e) {
    Logger.log("已付款信件讀取商品資料失敗: " + e.toString());
  }

  // ── 用 parseSummaryItems() 解析商品明細（與其他 email 一致）──────
  var parsedItems = parseSummaryItems(itemsSummary ? itemsSummary.toString() : "");

  // ── 建立商品明細 HTML（HK 用 HKD，TW 用 TWD）────────────────────
  // 預建模糊比對用的 cleanKey 對照表
  var productCleanMap = {}; // cleanName -> { photo, price, twd, origName }
  for (var pm in productMap) {
    var cleanKey = pm.toLowerCase().replace(/[^a-z0-9一-龥]/g, "");
    productCleanMap[cleanKey] = { photo: productMap[pm].photo, price: productMap[pm].price, twd: productMap[pm].twd, origName: pm };
  }

  function findProductData(name) {
    if (!name) return null;
    // 精確比對
    if (productMap[name]) return productMap[name];
    // 清理後精確比對
    var cleanName = name.toLowerCase().replace(/[^a-z0-9一-龥]/g, "");
    if (productCleanMap[cleanName]) return productCleanMap[cleanName];
    // 模糊包含比對
    for (var ck in productCleanMap) {
      if (cleanName && ck && (cleanName.indexOf(ck) !== -1 || ck.indexOf(cleanName) !== -1)) {
        return productCleanMap[ck];
      }
    }
    return null;
  }

  var fallbackImg = "data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%221%22%20height%3D%221%22%3E%3Crect%20width%3D%221%22%20height%3D%221%22%20fill%3D%22%23111%22%2F%3E%3C%2Fsvg%3E";
  var displayCurrency = isTW ? "NT$" : "HK$";
  var displayColor = isTW ? "#2980b9" : "#111111";

  var tableRowsHtml = "";
  for (var k = 0; k < parsedItems.length; k++) {
    var pItem = parsedItems[k];
    var nameWithStyle = pItem.name + (pItem.sub ? " (" + pItem.sub + ")" : "");
    var quantity      = pItem.qty || 1;
    var matched    = findProductData(pItem.name);
    var imgUrl     = (matched && matched.photo && matched.photo.indexOf("http") === 0) ? matched.photo : fallbackImg;
    var hkdPrice   = matched ? (matched.price || 0) : 0;
    var twdPrice   = matched ? (matched.twd   || 0) : 0;
    var displayPrice = isTW ? twdPrice : hkdPrice;
    var itemSubtotal = displayPrice * quantity;
    tableRowsHtml += "<tr style=\"border-bottom: 1px solid #eeeeee;\">" +
      "<td style=\"padding: 12px 10px; font-size: 14px; color: #333333; padding-left: 15px;\">" +
        "<table style=\"width: 100%; border-collapse: collapse; border: none;\"><tr>" +
          "<td style=\"padding: 0; width: 50px; vertical-align: middle;\">" +
            "<img src=\"" + imgUrl + "\" style=\"width: 50px; height: 50px; object-fit: cover; border-radius: 4px; border: 1px solid #eef0ed; display: block;\" />" +
          "</td>" +
          "<td style=\"padding: 0; padding-left: 12px; vertical-align: middle; line-height: 1.4;\">" + nameWithStyle + "</td>" +
        "</tr></table>" +
      "</td>" +
      "<td style=\"padding: 12px 10px; font-size: 14px; color: #555555; text-align: center;\">" + quantity + "</td>" +
      "<td style=\"padding: 12px 10px; font-size: 14px; color: " + displayColor + "; text-align: right; font-weight: 500; padding-right: 15px; width: 80px;\">" +
        displayCurrency + " " + (itemSubtotal > 0 ? itemSubtotal : "\u2014") +
      "</td></tr>";
  }

  // ════════════════════════════════════════════════════════════════
  // 🇹🇼 台灣版已付款 email
  // ════════════════════════════════════════════════════════════════
  if (isTW) {
    var twEmailSubject = `[已付款] 代購 ${currentEventName} — 886tw.81jp 台灣`;
    var twEmailBody = `
      <div style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width:600px; margin:0 auto; color:#333333; line-height:1.6; padding:20px; background-color:#f7f6f3;">
        <h2 style="font-size:18px; font-weight:500; border-bottom:1px solid #e5e5e5; padding-bottom:10px; color:#1a1a1a;">代購 ${currentEventName}</h2>
        <p style="font-size:14px; color:#1565c0; font-weight:500;">我們已收到客人的付款，商品運送回台時，會再次以 Email 通知客人於賣貨便下單，謝謝！</p>
        <p style="font-size:14px; background:#f2f2f2; padding:8px 12px; border-radius:4px; display:inline-block;"><strong>訂單編號：</strong> <span style="color:#111111;">${orderId}</span></p>

        <!-- 進度條（已付款亮起藍燈） -->
        <div style="margin:35px 0; padding:0 10px;">
          <div style="position:relative; display:flex; justify-content:space-between; align-items:center; width:100%;">
            <div style="position:absolute; top:5px; left:0; right:0; height:1px; background-color:#e5e5e5; z-index:1;"></div>
            <div style="position:relative; z-index:2; text-align:center; width:25%;">
              <div style="width:11px; height:11px; background-color:#2980b9; border-radius:50%; margin:0 auto; border:2px solid #ffffff; box-shadow:0 0 0 1px #2980b9;"></div>
              <div style="font-size:12px; font-weight:bold; color:#2980b9; margin-top:8px;">已收到訂單</div>
            </div>
            <div style="position:relative; z-index:2; text-align:center; width:25%;">
              <div style="width:11px; height:11px; background-color:#2980b9; border-radius:50%; margin:0 auto; border:2px solid #ffffff; box-shadow:0 0 0 1px #2980b9;"></div>
              <div style="font-size:12px; font-weight:bold; color:#2980b9; margin-top:8px;">已付款</div>
            </div>
            <div style="position:relative; z-index:2; text-align:center; width:25%;">
              <div style="width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;"></div>
              <div style="font-size:11px; color:#888888; margin-top:8px;">已到台</div>
            </div>
            <div style="position:relative; z-index:2; text-align:center; width:25%;">
              <div style="width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;"></div>
              <div style="font-size:11px; color:#888888; margin-top:8px;">已完成取貨</div>
            </div>
          </div>
        </div>

        <h3 style="font-size:14px; font-weight:bold; margin-top:25px; margin-bottom:10px; color:#444444;">📋 訂購商品明細</h3>
        <table style="width:100%; border-collapse:collapse; background:#ffffff; border:1px solid #eef0ed; border-radius:6px; overflow:hidden;">
          <thead>
            <tr style="background-color:#f0f7ff; border-bottom:1px solid #eef0ed;">
              <th style="padding:10px 15px; text-align:left; font-size:13px; color:#666666; font-weight:500;">商品名稱與款式</th>
              <th style="padding:10px; text-align:center; font-size:13px; color:#666666; font-weight:500; width:60px;">數量</th>
              <th style="padding:10px 15px; text-align:right; font-size:13px; color:#666666; font-weight:500; width:90px;">小計</th>
            </tr>
          </thead>
          <tbody>${tableRowsHtml}</tbody>
        </table>

        <p style="text-align:right; font-size:16px; font-weight:bold; margin-top:15px; color:#1a1a1a;">
          商品總金額：<span style="font-size:20px; color:#2980b9;">NT$ ${totalAmount}</span>
        </p>
        <p style="text-align:right; font-size:12px; color:#888; margin-top:4px;">
          ＋ 國際運費（NT$20／50g）＋ NT$38 賣貨便運費（出貨時另行通知）
        </p>

        <div style="margin-top:30px; padding-top:15px; border-top:1px dashed #dddddd; font-size:13px; color:#777777; line-height:1.8;">
          <p style="margin:5px 0;"></p>
          <p style="margin:5px 0;">📦 貨品到台及寄出後，也都會有 Email 通知您。</p>
          <p style="margin:5px 0;">如有任何問題，都可以直接回覆此 Email 或到 Threads inbox 我們，並提供訂單編號查詢，謝謝！</p>
        </div>
        <div style="margin-top:40px; font-size:13px; color:#333333; line-height:1.5;">
          <strong>886tw.81jp</strong><br>
          <a href="https://www.threads.com/@886tw.81jp?igshid=NTc4MTIwNjQ2YQ==" style="color:#2980b9;">Threads @886tw.81jp</a>　<a href="https://www.instagram.com/886tw.81jp?igsh=MW8zMmVncGVwNmd1dg%3D%3D&utm_source=qr" style="color:#e1306c;">Instagram @886tw.81jp</a>
        </div>
      </div>`;

    if (targetEmail !== "" && targetEmail.indexOf("@") !== -1) {
      MailApp.sendEmail({
        to: targetEmail,
        subject: twEmailSubject,
        body: "我們已收到您的付款，訂單編號：" + orderId + "。請查看 HTML 郵件以獲取最新訂單進度。",
        htmlBody: twEmailBody
      });
    } else {
      Logger.log("無法發送 TW 已付款郵件，無效的 Email: " + targetEmail);
    }

  // ════════════════════════════════════════════════════════════════
  // 🇭🇰 香港版已付款 email（原有邏輯完整保留）
  // ════════════════════════════════════════════════════════════════
  } else {
    var emailSubject = `[已付款] 代購 ${currentEventName} - 852hk.81jp`;
    var emailBodyHtml = `
      <div style="font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width: 600px; margin: 0 auto; color: #333333; line-height: 1.6; padding: 20px; background-color: #fbfbfa;">
        <h2 style="font-size: 18px; font-weight: 500; border-bottom: 1px solid #e5e5e5; padding-bottom: 10px; color: #111111;">代購 ${currentEventName}</h2>
        <p style="font-size: 14px; color: #1a73e8; font-weight: 500;">我哋已收到客人嘅付款，商品到港途中時，我哋會再次通知客人，謝謝！</p>
        <p style="font-size: 14px; background: #f2f2f2; padding: 8px 12px; border-radius: 4px; display: inline-block;"><strong>訂單編號：</strong> <span style="color: #111111;">${orderId}</span></p>

        <div style="margin: 35px 0; padding: 0 10px;">
          <div style="position: relative; display: flex; justify-content: space-between; align-items: center; width: 100%;">
            <div style="position: absolute; top: 5px; left: 0; right: 0; height: 1px; background-color: #e5e5e5; z-index: 1;"></div>
            <div style="position: relative; z-index: 2; text-align: center; width: 25%;">
              <div style="width: 11px; height: 11px; background-color: #111111; border-radius: 50%; margin: 0 auto; border: 2px solid #ffffff; box-shadow: 0 0 0 1px #111111;"></div>
              <div style="font-size: 12px; font-weight: bold; color: #111111; margin-top: 8px;">已收到訂單</div>
            </div>
            <div style="position: relative; z-index: 2; text-align: center; width: 25%;">
              <div style="width: 11px; height: 11px; background-color: #111111; border-radius: 50%; margin: 0 auto; border: 2px solid #ffffff; box-shadow: 0 0 0 1px #111111;"></div>
              <div style="font-size: 12px; font-weight: bold; color: #111111; margin-top: 8px;">已付款</div>
            </div>
            <div style="position: relative; z-index: 2; text-align: center; width: 25%;">
              <div style="width: 11px; height: 11px; background-color: #d1d5db; border-radius: 50%; margin: 0 auto; border: 2px solid #ffffff;"></div>
              <div style="font-size: 11px; color: #888888; margin-top: 8px;">到港途中</div>
            </div>
            <div style="position: relative; z-index: 2; text-align: center; width: 25%;">
              <div style="width: 11px; height: 11px; background-color: #d1d5db; border-radius: 50%; margin: 0 auto; border: 2px solid #ffffff;"></div>
              <div style="font-size: 11px; color: #888888; margin-top: 8px;">已寄出</div>
            </div>
          </div>
        </div>

        <h3 style="font-size: 14px; font-weight: bold; margin-top: 25px; margin-bottom: 10px; color: #444444;">📋 訂購商品明細</h3>
        <table style="width: 100%; border-collapse: collapse; background: #ffffff; border: 1px solid #eef0ed; border-radius: 6px; overflow: hidden;">
          <thead>
            <tr style="background-color: #fafbfa; border-bottom: 1px solid #eef0ed;">
              <th style="padding: 10px; text-align: left; font-size: 13px; color: #666666; font-weight: 500; padding-left: 15px;">商品名稱與款式</th>
              <th style="padding: 10px; text-align: center; font-size: 13px; color: #666666; font-weight: 500; width: 60px;">數量</th>
              <th style="padding: 10px; text-align: right; font-size: 13px; color: #666666; font-weight: 500; width: 80px; padding-right: 15px;">小計</th>
            </tr>
          </thead>
          <tbody>${tableRowsHtml}</tbody>
        </table>

        <p style="text-align: right; font-size: 16px; font-weight: bold; margin-top: 15px; color: #111111;">
          總金額 (含郵費)：<span style="font-size: 20px; color: #ff4d4d;">HK$ ${totalAmount}</span>
        </p>

        <h3 style="font-size: 14px; font-weight: bold; margin-top: 25px; margin-bottom: 5px; color: #444444;">📍 易寄取收件資料</h3>
        <div style="background: #ffffff; border: 1px solid #eef0ed; padding: 15px; border-radius: 6px; font-size: 13px; color: #555555; line-height: 1.8;">
          <strong>收件人姓名：</strong> ${recipientName || "-"}<br>
          <strong>聯絡電話：</strong> ${recipientPhone || "-"}<br>
          <strong>自提點類別：</strong> ${pickupType || "-"}<br>
          <strong>自提點編號：</strong> ${pickupCode || "-"}<br>
          <strong>自提點名稱：</strong> ${pickupName || "-"}<br>
          <strong>自提點地址：</strong> ${pickupAddress || "-"}
        </div>

        <div style="margin-top: 30px; padding-top: 15px; border-top: 1px dashed #dddddd; font-size: 13px; color: #777777; line-height: 1.7;">
          <p style="margin: 5px 0;">✨ 確認付款後，你會再收到一封訂單確認 Email。</p>
          <p style="margin: 5px 0;">📦 貨品到貨及寄出後，也都會有 Email 通知你。</p>
          <p style="margin: 5px 0;">如有任何問題，都可以直接回覆此 Email 或到 Threads inbox 我哋，並提供訂單編號查詢，謝謝！</p>
        </div>
        <div style="margin-top: 40px; font-size: 13px; color: #333333; line-height: 1.5;">
          <strong>852hk.81jp</strong><br>
          Threads: <a href="https://www.threads.net/@852hk.81jp" target="_blank" style="color: #111111; text-decoration: underline;">https://www.threads.net/@852hk.81jp</a>
        </div>
      </div>`;

    if (targetEmail !== "" && targetEmail.indexOf("@") !== -1) {
      MailApp.sendEmail({
        to: targetEmail,
        subject: emailSubject,
        body: "我哋已收到客人嘅付款，您的訂單編號為: " + orderId + "。請查看 HTML 郵件內容以獲得最新訂單進度。",
        htmlBody: emailBodyHtml
      });
    } else {
      Logger.log("無法發送已付款郵件，無效的 Email: " + targetEmail);
    }
  }
}

// =================================================================
// syncAndGetDeadlines：自動同步 event 分頁到「收單截止時間」
// =================================================================
function syncAndGetDeadlines(ss) {
  var EXCLUDED = ["訂單紀錄", "易寄取地址", "Blank", "收單截止時間", "購貨紀錄", "盈利紀錄"];
  var allSheets = ss.getSheets();
  var deadlineSheet = ss.getSheetByName("收單截止時間");
  
  if (!deadlineSheet) {
    deadlineSheet = ss.insertSheet("收單截止時間");
    deadlineSheet.appendRow(["Event", "截止時間 (YYYY/MM/DD HH:MM)"]);
  }
  
  var existingRows = deadlineSheet.getDataRange().getValues();
  var existingEvents = existingRows.slice(1).map(function(r) { return r[0].toString(); });
  
  for (var i = 0; i < allSheets.length; i++) {
    var name = allSheets[i].getName();
    if (EXCLUDED.indexOf(name) === -1 && existingEvents.indexOf(name) === -1) {
      deadlineSheet.appendRow([name, ""]);
      existingEvents.push(name);
    }
  }
  
  var finalRows = deadlineSheet.getDataRange().getValues();
  var result = {};
  for (var r = 1; r < finalRows.length; r++) {
    result[finalRows[r][0].toString()] = finalRows[r][1] ? finalRows[r][1].toString() : "";
  }
  return result;
}

// =================================================================
// getOrderedQtyForEvent：統計某 event 各款式的已購總數（HK+TW合計）
// 回傳格式：{ "商品名|款式": 已購數量, ... }
// =================================================================
function getOrderedQtyForEvent(ss, eventName) {
  var cleanEventName = stripEndPrefix(eventName);
  var result = {};
  var orderSheet = ss.getSheetByName("訂單紀錄");
  if (!orderSheet) return result;

  var rows = orderSheet.getDataRange().getValues();

  for (var r = 1; r < rows.length; r++) {
    // C欄(index 2)：活動名稱；D欄(index 3)：itemsSummary
    if (!rows[r][0]) continue;
    var rowEvent = rows[r][2] ? rows[r][2].toString().trim() : "";
    if (rowEvent !== cleanEventName) continue;

    var summary = rows[r][3] ? rows[r][3].toString() : "";
    if (!summary) continue;

    // 解析 itemsSummary：格式為 "商品A (款式1 x2, 款式2 x1), 商品B (款式3 x1)"
    var items = [];
    var currentItem = "";
    var inBrackets = false;
    for (var c = 0; c < summary.length; c++) {
      var ch = summary[c];
      if (ch === '(' || ch === '（') { inBrackets = true; currentItem += ch; }
      else if (ch === ')' || ch === '）') { inBrackets = false; currentItem += ch; }
      else if ((ch === ',' || ch === '，') && !inBrackets) {
        if (currentItem.trim()) items.push(currentItem.trim());
        currentItem = "";
      } else { currentItem += ch; }
    }
    if (currentItem.trim()) items.push(currentItem.trim());

    for (var i = 0; i < items.length; i++) {
      var itemText = items[i].trim();
      if (!itemText) continue;

      // 先去掉結尾的全局數量（如 "商品名 x2"）
      var globalQty = 1;
      var tailMatch = itemText.match(/[xX]\s*(\d+)\s*$/);
      if (tailMatch) {
        globalQty = parseInt(tailMatch[1]) || 1;
        itemText = itemText.replace(/[xX]\s*\d+\s*$/, "").trim();
      }

      var productName = "";
      var subsText = "";

      // 解析括號內的款式
      var openIdx = itemText.indexOf('(');
      var closeIdx = itemText.lastIndexOf(')');
      if (openIdx !== -1 && closeIdx !== -1 && closeIdx > openIdx) {
        productName = itemText.substring(0, openIdx).trim();
        subsText = itemText.substring(openIdx + 1, closeIdx).trim();
      } else {
        productName = itemText.trim();
      }

      if (!productName) continue;

      if (subsText) {
        // 括號內可能有多個款式：「款式1 x2, 款式2 x1」
        var subParts = subsText.split(/[,，]/);
        for (var s = 0; s < subParts.length; s++) {
          var part = subParts[s].trim();
          if (!part) continue;
          var subQty = globalQty;
          var subName = part;
          var subQtyMatch = part.match(/[xX]\s*(\d+)\s*$/);
          if (subQtyMatch) {
            subQty = parseInt(subQtyMatch[1]) || 1;
            subName = part.replace(/[xX]\s*\d+\s*$/, "").trim();
          }
          var key = productName + "|" + subName;
          result[key] = (result[key] || 0) + subQty;
        }
      } else {
        // 沒有款式的商品，key 用「商品名|」
        var key = productName + "|";
        result[key] = (result[key] || 0) + globalQty;
      }
    }
  }

  return result;
}

// =================================================================
// parseSubStockLimit：解析 J欄文字格式 "白色:1,黑色:0,藍色:5"
// 回傳 { "白色": 1, "黑色": 0, "藍色": 5 } 或 null（無限購）
// 相容舊格式純數字：回傳 { "__all__": n }
// =================================================================
function parseSubStockLimit(raw) {
  if (raw === null || raw === undefined || raw.toString().trim() === "") return null;
  var str = raw.toString().trim();
  // 舊格式：純數字
  if (/^\d+$/.test(str)) {
    return { "__all__": parseInt(str) };
  }
  var result = {};
  var parts = str.split(",");
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    if (!part) continue;
    var colonIdx = part.lastIndexOf(":");
    if (colonIdx === -1) continue;
    var subName = part.substring(0, colonIdx).trim();
    var subVal  = part.substring(colonIdx + 1).trim();
    result[subName] = (subVal === "" || isNaN(parseInt(subVal))) ? null : parseInt(subVal);
  }
  return Object.keys(result).length > 0 ? result : null;
}

// =================================================================
// serializeSubStockLimit：{ "白色": 1, "黑色": 0 } → "白色:1,黑色:0"
// =================================================================
function serializeSubStockLimit(obj) {
  if (!obj) return "";
  if (obj["__all__"] !== undefined) return obj["__all__"].toString();
  return Object.keys(obj).map(function(k) {
    return k + ":" + (obj[k] === null || obj[k] === undefined ? "" : obj[k]);
  }).join(",");
}

// =================================================================
// getOrderSummary：統計指定 event 的已付款訂單
// 回傳每個商品款式的需求數量、日幣總額、圖片、已購買數量
// =================================================================
function getOrderSummary(ss, eventName) {
  var orderSheet = ss.getSheetByName("訂單紀錄");
  var purchaseSheet = getPurchaseSheet(ss);
  var cleanEventName = stripEndPrefix(eventName);
  var productSheet = getSheetByEventName(ss, eventName);

  // ── 從商品分頁讀取圖片和日幣價格 ──
  var photoMap = {};   // { "商品名": url }
  var yenMap = {};     // { "商品名": yen }
  var pRows = [];
  if (productSheet) {
    pRows = productSheet.getDataRange().getValues();
    for (var pi = 1; pi < pRows.length; pi++) {
      var pn = pRows[pi][0] ? pRows[pi][0].toString().trim() : "";
      if (!pn) continue;
      photoMap[pn] = pRows[pi][5] ? pRows[pi][5].toString().trim() : "";
      yenMap[pn]   = parseFloat(pRows[pi][2]) || 0;
    }
  }

  // ── 從訂單紀錄統計已付款訂單 ──
  var summaryMap = {}; // key = "商品名|款式" → { name, sub, paidQty, unpaidQty, totalYen, photo, yenEach }
  if (orderSheet) {
    var oRows = orderSheet.getDataRange().getValues();
    for (var r = 1; r < oRows.length; r++) {
      if (!oRows[r][0]) continue;
      var rowEvent  = oRows[r][2] ? oRows[r][2].toString().trim() : "";
      var rowStatus = oRows[r][15] ? oRows[r][15].toString().trim() : "";
      if (rowEvent !== cleanEventName) continue;
      if (rowStatus === "已取消") continue; // 排除已取消
      var isPaid = (rowStatus === "已收款" || rowStatus === "已付款");

      var summary = oRows[r][3] ? oRows[r][3].toString() : "";
      var parsed = parseSummaryItems(summary);

      for (var k = 0; k < parsed.length; k++) {
        var item = parsed[k];
        var key = item.name + "|" + item.sub;
        if (!summaryMap[key]) {
          var photo = photoMap[item.name] || "";
          var yen   = yenMap[item.name] || 0;
          if (!photo || yen === 0) {
            var cleanItemName = item.name.toLowerCase().replace(/[\s\u3000]/g, "");
            for (var pm in photoMap) {
              var cleanPm = pm.toLowerCase().replace(/[\s\u3000]/g, "");
              if (cleanItemName === cleanPm || cleanItemName.indexOf(cleanPm) !== -1 || cleanPm.indexOf(cleanItemName) !== -1) {
                if (!photo) photo = photoMap[pm];
                if (yen === 0) yen = yenMap[pm] || 0;
                break;
              }
            }
          }
          summaryMap[key] = { name: item.name, sub: item.sub, paidQty: 0, unpaidQty: 0, totalYen: 0, photo: photo, yenEach: yen };
        }
        if (isPaid) {
          summaryMap[key].paidQty  += item.qty;
          summaryMap[key].totalYen += (summaryMap[key].yenEach * item.qty);
        } else {
          summaryMap[key].unpaidQty += item.qty;
        }
      }
    }
  }

  // ── 從購貨紀錄讀取已購買數量 ──
  // ── 從購貨紀錄讀取（優先 per-event sheet）──
  var purchasedMap = {};
  var settledMap   = {};
  var perEvSheet = getEventPurchaseSheet(ss, eventName, false);
  if (perEvSheet) {
    var peRows = perEvSheet.getDataRange().getValues();
    for (var pr = 1; pr < peRows.length; pr++) {
      if (!peRows[pr][0]) continue;
      var prName   = peRows[pr][0] ? peRows[pr][0].toString().trim() : "";
      var prSub    = peRows[pr][1] ? peRows[pr][1].toString().trim() : "";
      var prBought = Number(peRows[pr][5]) || 0; // F(index5)=已購買
      var prSettld = Number(peRows[pr][6]) || 0; // G(index6)=已結算
      purchasedMap[prName + "|" + prSub] = prBought;
      settledMap[prName + "|" + prSub]   = prSettld;
    }
  } else if (purchaseSheet) {
    var purchaseRows = purchaseSheet.getDataRange().getValues();
    for (var pr2 = 1; pr2 < purchaseRows.length; pr2++) {
      var prEv = purchaseRows[pr2][0] ? purchaseRows[pr2][0].toString().trim() : "";
      if (!prEv || prEv !== cleanEventName) continue;
      var prNm2 = purchaseRows[pr2][1] ? purchaseRows[pr2][1].toString().trim() : "";
      var prSb2 = purchaseRows[pr2][2] ? purchaseRows[pr2][2].toString().trim() : "";
      if (!prNm2) continue;
      purchasedMap[prNm2 + "|" + prSb2] = Number(purchaseRows[pr2][6]) || 0;
      settledMap[prNm2 + "|" + prSb2]   = Number(purchaseRows[pr2][7]) || 0;
    }
  }

  // ── 整合回傳 ──
  var result = [];
  // 先按商品名稱分組
  var productGroups = {}; // { "商品名": [ { sub, qty, totalYen, photo, yenEach, purchased } ] }
  for (var key in summaryMap) {
    var item = summaryMap[key];
    if (!productGroups[item.name]) productGroups[item.name] = [];
    var purchasedQty = purchasedMap[key] || 0;
    var remaining    = Math.max(0, item.qty - purchasedQty);
    // 從 event 分頁讀取 HKD 和 TWD 單價
    var priceHKD = 0, priceTWD = 0;
    for (var pp = 1; pp < pRows.length; pp++) {
      var ppName = pRows[pp][0] ? pRows[pp][0].toString().trim() : "";
      var cleanItem = item.name.toLowerCase().replace(/[\s\u3000]/g, "");
      var cleanPP   = ppName.toLowerCase().replace(/[\s\u3000]/g, "");
      if (ppName === item.name || (cleanItem && cleanPP && (cleanItem === cleanPP || cleanItem.indexOf(cleanPP) !== -1 || cleanPP.indexOf(cleanItem) !== -1))) {
        priceHKD = parseFloat(pRows[pp][4]) || 0; // E欄 HKD
        priceTWD = parseFloat(pRows[pp][8]) || 0; // I欄 TWD
        break;
      }
    }

    var purchasedQty = purchasedMap[key] || 0;
    var paidQty      = item.paidQty || 0;
    var unpaidQty    = item.unpaidQty || 0;
    var allQty       = paidQty + unpaidQty;
    var remaining    = Math.max(0, paidQty - purchasedQty);
    productGroups[item.name].push({
      sub:        item.sub,
      qty:        paidQty,         // 已付款訂單需求（向下相容）
      paidQty:    paidQty,
      unpaidQty:  unpaidQty,
      allQty:     allQty,
      totalYen:   item.totalYen,
      photo:      item.photo,
      yenEach:    item.yenEach,
      price:      priceHKD,
      twd:        priceTWD,
      purchased:  purchasedQty,
      settled:    settledMap[key] || 0,
      remaining:  remaining,
      isComplete: purchasedQty >= paidQty && paidQty > 0
    });
  }
  // ── 批次購買歷史（H欄=index7 開始），同日合併 ──
  var batchMap = {};
  var batchDates = [];
  var batchSheet = perEvSheet || getPurchaseSheet(ss); // 優先 per-event
  Logger.log("[getOrderSummary] eventName=" + eventName + " cleanName=" + cleanEventName);
  Logger.log("[getOrderSummary] perEvSheet=" + (perEvSheet ? perEvSheet.getName() : "null"));
  Logger.log("[getOrderSummary] batchSheet=" + (batchSheet ? batchSheet.getName() : "null"));
  if (batchSheet) {
    var bRows = batchSheet.getDataRange().getValues();
    var bNumCols = batchSheet.getLastColumn();
    var bNameIdx = perEvSheet ? 0 : 1; // per-event: A=名稱; 舊版: B=名稱
    var bSubIdx  = perEvSheet ? 1 : 2;
    Logger.log("[getOrderSummary] bNumCols=" + bNumCols + " row1=" + JSON.stringify(bRows[0]));
    for (var bc = 7; bc < bNumCols; bc++) { // per-event 從 H=index7 開始
      var bDateRaw = bRows[0][bc];
      var bDate = "";
      if (bDateRaw instanceof Date) {
        // Google Sheet 自動把日期字串轉成 Date 物件，需重新格式化
        bDate = Utilities.formatDate(bDateRaw, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else if (bDateRaw) {
        bDate = bDateRaw.toString().trim();
        // ISO 格式轉換 (e.g. "2026-05-30T16:00:00.000Z")
        if (bDate.indexOf("T") !== -1) {
          try { bDate = Utilities.formatDate(new Date(bDate), Session.getScriptTimeZone(), "yyyy/MM/dd"); } catch(e) {}
        }
      }
      // 跳過非日期欄（如 SubTotal、空白等）
      if (!bDate || bDate === "SubTotal" || bDate.length < 6) continue;
      if (!batchMap[bDate]) { batchMap[bDate] = {}; batchDates.push(bDate); }
      for (var br = 1; br < bRows.length; br++) {
        if (!perEvSheet) {
          var brEv = bRows[br][0] ? bRows[br][0].toString().trim() : "";
          if (brEv !== cleanEventName) continue;
        }
        var brNm = bRows[br][bNameIdx] ? bRows[br][bNameIdx].toString().trim() : "";
        var brSb = bRows[br][bSubIdx]  ? bRows[br][bSubIdx].toString().trim()  : "";
        var brQ  = Number(bRows[br][bc]) || 0;
        if (brQ > 0) {
          var bKey = brNm + "|" + brSb;
          batchMap[bDate][bKey] = (batchMap[bDate][bKey] || 0) + brQ;
        }
      }
    }
  }
  // 轉成陣列（只保留有資料的日期）
  var batches = [];
  for (var di = 0; di < batchDates.length; di++) {
    var bd = batchDates[di];
    if (Object.keys(batchMap[bd]).length > 0) {
      batches.push({ date: bd, items: batchMap[bd] });
    }
  }

  for (var pName in productGroups) {
    result.push({ name: pName, subs: productGroups[pName] });
  }
  return { items: result, batches: batches };
}

// =================================================================
// getOrderSummaryAll：逐個活動計「要買幾多錢」同「買咗幾多錢」，
// 俾後台「購貨總覽」一眼睇晒邊個團仲未買完。
// =================================================================
function getOrderSummaryAll(ss) {
  var sysSheets = ["訂單紀錄", "易寄取地址", "Blank", "收單截止時間", "購貨紀錄", "盈利紀錄", "Nissen精選"];
  var sheets = ss.getSheets();
  var events = [];

  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if (sheetName.startsWith("[Data]")) continue;
    if (sysSheets.indexOf(sheetName) !== -1) continue;

    var needYen = 0, boughtYen = 0, productCount = 0;
    try {
      var items = (getOrderSummary(ss, sheetName).items) || [];
      productCount = items.length;
      for (var it = 0; it < items.length; it++) {
        var subs = items[it].subs || [];
        for (var s = 0; s < subs.length; s++) {
          var yenEach = subs[s].yenEach || 0;
          needYen   += yenEach * (subs[s].paidQty   || 0);
          boughtYen += yenEach * (subs[s].purchased || 0);
        }
      }
    } catch (evErr) {
      // 一個活動出事唔應該拖冧成張表，記低就算
      Logger.log("getOrderSummaryAll 略過「" + sheetName + "」：" + evErr);
      continue;
    }

    // 完全冇訂單嘅活動唔使佔行
    if (productCount === 0) continue;

    events.push({
      name:         sheetName,
      displayName:  stripEndPrefix(sheetName),
      productCount: productCount,
      needYen:      Math.round(needYen),
      boughtYen:    Math.round(boughtYen)
    });
  }

  return events;
}

// =================================================================
// parseSummaryItems：解析 itemsSummary 字串成 [{name, sub, qty}] 陣列
// 支援格式：
//   商品名 ( x1)                          → 款式空，qty=1
//   商品名(含括號) (款式 x1)               → 最後括號才是款式+數量
//   商品名 (款式A x2, 款式B x1)            → 多款式分行
//   商品名1 ( x1), 商品名2 ( x1)           → 多商品逗號分隔
// =================================================================
function parseSummaryItems(summaryText) {
  if (!summaryText) return [];
  var items = [];

  // Step 1：按最外層逗號分割（所有括號類型 () {} [] 都計入深度）
  var parts = [];
  var current = "";
  var depth = 0;
  for (var c = 0; c < summaryText.length; c++) {
    var ch = summaryText[c];
    if (ch === "(" || ch === "\uff08" || ch === "{" || ch === "[") { depth++; current += ch; }
    else if (ch === ")" || ch === "\uff09" || ch === "}" || ch === "]") { depth--; current += ch; }
    else if ((ch === "," || ch === "\uff0c") && depth === 0) {
      if (current.trim()) parts.push(current.trim());
      current = "";
    } else { current += ch; }
  }
  if (current.trim()) parts.push(current.trim());

  // 尺碼判斷（只有這些才視為款式，與前端一致）
  var SIZE_PATTERN = /^(XS|S|M|L|XL|XXL|2XL|3XL|XXXL|4XL|XXXXL|Free Size|One Size|FREE|OS|F\/S)$/i;

  // Step 2：解析每個 part（支援新舊兩種格式）
  for (var i = 0; i < parts.length; i++) {
    var part = parts[i].trim();
    if (!part) continue;
    var qty = 1;
    var sub = "";

    // ══ 新格式：ProductName {Sub} [xN] ══
    // e.g. "ZUTOMAYO Tee(Vintage Black) {XXL} [x1]"
    var newFull = part.match(/\{([^}]*)\}\s*\[x(\d+)\]\s*$/);
    if (newFull) {
      sub = newFull[1].trim();
      qty = parseInt(newFull[2]) || 1;
      part = part.slice(0, part.lastIndexOf("{")).trim();
      items.push({ name: part, sub: sub, qty: qty });
      continue;
    }
    // e.g. "ProductName [x2]" — 無款式
    var newQtyOnly = part.match(/\[x(\d+)\]\s*$/);
    if (newQtyOnly) {
      qty = parseInt(newQtyOnly[1]) || 1;
      part = part.slice(0, part.lastIndexOf("[x")).trim();
      items.push({ name: part, sub: "", qty: qty });
      continue;
    }

    // ══ 舊格式（backward compatibility）══
    // ── 格式 A：末尾 (SIZE xN) ──
    var fmtA = part.match(/\(([^)]+?)\s+[xX\xd7]\s*(\d+)\s*\)\s*$/);
    if (fmtA && SIZE_PATTERN.test(fmtA[1].trim())) {
      sub = fmtA[1].trim();
      qty = parseInt(fmtA[2]) || 1;
      part = part.slice(0, part.lastIndexOf(fmtA[0])).trim();
      items.push({ name: part, sub: sub, qty: qty });
      continue;
    }
    // ── 格式 B：末尾 ( xN) ──
    var fmtB = part.match(/\(\s*[xX\xd7]\s*(\d+)\s*\)\s*$/);
    if (fmtB) {
      qty = parseInt(fmtB[1]) || 1;
      part = part.slice(0, part.lastIndexOf(fmtB[0])).trim();
    }
    // ── 剩餘末尾括號是否為尺碼？ ──
    var lo = part.lastIndexOf("(");
    var lc = part.lastIndexOf(")");
    if (lo !== -1 && lc > lo && lc === part.length - 1) {
      var inside = part.substring(lo + 1, lc).trim();
      if (SIZE_PATTERN.test(inside)) {
        sub = inside;
        part = part.substring(0, lo).trim();
      }
    }

    items.push({ name: part, sub: sub, qty: qty });
  }
  return items;
}

// =================================================================
// savePurchased：批量寫回「購貨紀錄」E欄
// payload: { eventName, items: [{name, sub, qty, purchased}] }
// =================================================================
function savePurchased(ss, eventName, items) {
  // 優先使用 per-event sheet
  var sheet = getEventPurchaseSheet(ss, eventName, true); // 不存在就建立
  if (!sheet) return;

  var rows = sheet.getDataRange().getValues();
  var numCols = sheet.getLastColumn();

  // 找下一個可用批次欄（從 H=col8 開始）
  var batchCol = 8;
  for (var c = 8; c <= numCols; c++) {
    var hv = rows[0] && rows[0][c-1] ? rows[0][c-1].toString().trim() : "";
    if (!hv) { batchCol = c; break; }
    if (c === numCols) batchCol = numCols + 1;
  }

  // 寫入批次日期到第1行（強制文字格式，避免 Google Sheet 自動轉成 Date 物件）
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  sheet.getRange(1, batchCol).setNumberFormat("@").setValue(today);

  // 從 event sheet 取日幣單價
  var yenMap = {};
  var eventSheet = getSheetByEventName(ss, eventName);
  if (eventSheet) {
    var eRows = eventSheet.getDataRange().getValues();
    for (var er = 1; er < eRows.length; er++) {
      var en = eRows[er][0] ? eRows[er][0].toString().trim() : "";
      if (en) yenMap[en] = parseFloat(eRows[er][2]) || 0;
    }
  }

  // 從訂單紀錄統計 paidQty 和 unpaidQty
  var paidMap = {}, unpaidMap = {};
  var cleanEventName = stripEndPrefix(eventName);
  var orderSheet = ss.getSheetByName("訂單紀錄");
  if (orderSheet) {
    var oRows = orderSheet.getDataRange().getValues();
    for (var oi = 1; oi < oRows.length; oi++) {
      if (!oRows[oi][0]) continue;
      var rowEvent  = oRows[oi][2] ? oRows[oi][2].toString().trim() : "";
      var rowStatus = oRows[oi][15] ? oRows[oi][15].toString().trim() : "";
      if (rowEvent !== cleanEventName) continue;
      if (rowStatus === "已取消") continue;
      var isPaid = (rowStatus === "已收款" || rowStatus === "已付款");
      var parsed = parseSummaryItems(oRows[oi][3] ? oRows[oi][3].toString() : "");
      for (var pk = 0; pk < parsed.length; pk++) {
        var pkey = parsed[pk].name + "|" + parsed[pk].sub;
        if (isPaid) paidMap[pkey] = (paidMap[pkey] || 0) + parsed[pk].qty;
        else unpaidMap[pkey] = (unpaidMap[pkey] || 0) + parsed[pk].qty;
      }
    }
  }

  for (var k = 0; k < items.length; k++) {
    var item = items[k];
    if (!item.purchased || item.purchased <= 0) continue;
    var itemSub  = item.sub  ? item.sub.toString().trim()  : "";
    var itemName = item.name ? item.name.toString().trim() : "";
    var mapKey = itemName + "|" + itemSub;
    var found = false;

    for (var r = 1; r < rows.length; r++) {
      if (!rows[r][0]) continue;
      var rName = rows[r][0].toString().trim(); // A欄
      var rSub  = rows[r][1] ? rows[r][1].toString().trim() : ""; // B欄
      if (rName === itemName && rSub === itemSub) {
        var curBought = Number(rows[r][5]) || 0; // F欄(index5)=已購買
        sheet.getRange(r + 1, 6).setValue(curBought + item.purchased); // F=col6
        sheet.getRange(r + 1, batchCol).setValue(item.purchased);       // 批次欄
        rows[r][5] = curBought + item.purchased;
        // 補齊日幣單價（若 C 欄為空）
        if (!rows[r][2] && yenMap[itemName]) {
          sheet.getRange(r + 1, 3).setValue(yenMap[itemName]);
          rows[r][2] = yenMap[itemName];
        }
        found = true;
        break;
      }
    }
    if (!found) {
      // 新增行：A=名稱, B=款式, C=日幣, D=已付款, E=未付款, F=已購買, G=已結算
      var yenEach   = yenMap[itemName]   || 0;
      var paidQty   = paidMap[mapKey]    || 0;
      var unpaidQty = unpaidMap[mapKey]  || 0;
      var newRow = [itemName, itemSub, yenEach, paidQty, unpaidQty, item.purchased, ""];
      while (newRow.length < batchCol - 1) newRow.push("");
      newRow.push(item.purchased);
      sheet.appendRow(newRow);
      rows.push(newRow);
    }
  }
  // ── 即時同步買貨區到 Supabase ──
  try { syncPurchaseRecordsToSupabase(eventName); } catch(spErr) { Logger.log("Supabase purchase sync err: " + spErr); }
}

// =================================================================
// 共用：建立取消 email 的商品明細表格 HTML（與其他 email 一致風格）
// =================================================================
function buildCancelItemsTable(items, productSheet, isTW) {
  var fallbackImg = "data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%221%22%20height%3D%221%22%3E%3Crect%20width%3D%221%22%20height%3D%221%22%20fill%3D%22%23111%22%2F%3E%3C%2Fsvg%3E";
  var currency  = isTW ? "NT$" : "HK$";
  var priceCol  = isTW ? 8 : 4; // I欄TWD=8, E欄HKD=4

  // 從商品分頁建立 photo 和 price 對照表
  var photoMap = {}, priceMap = {};
  if (productSheet) {
    var pRows = productSheet.getDataRange().getValues();
    for (var p = 1; p < pRows.length; p++) {
      var pn = pRows[p][0] ? pRows[p][0].toString().trim() : "";
      if (!pn) continue;
      photoMap[pn] = pRows[p][5] ? pRows[p][5].toString().trim() : "";
      priceMap[pn] = parseFloat(pRows[p][priceCol]) || 0;
    }
  }

  function findPhoto(name) {
    if (photoMap[name]) return photoMap[name];
    var clean = name.toLowerCase().replace(/[\s　]/g, "");
    for (var k in photoMap) {
      var kc = k.toLowerCase().replace(/[\s　]/g, "");
      if (clean === kc || clean.indexOf(kc) !== -1 || kc.indexOf(clean) !== -1) return photoMap[k] || "";
    }
    return "";
  }

  var html = "";
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var displayName = item.name + (item.sub ? " (" + item.sub + ")" : "");
    var qty = item.qty || 1;
    var photo = findPhoto(item.name);
    var imgUrl = (photo && photo.indexOf("http") === 0) ? photo : fallbackImg;
    var price = priceMap[item.name] || 0;
    var subtotal = price * qty;

    html += '<tr style="border-bottom:1px solid #eeeeee;">' +
      '<td style="padding:12px 10px 12px 15px; font-size:14px; color:#333333;">' +
        '<table style="width:100%; border-collapse:collapse; border:none;"><tr>' +
          '<td style="padding:0; width:50px; vertical-align:middle;">' +
            '<img src="' + imgUrl + '" style="width:50px; height:50px; object-fit:cover; border-radius:4px; border:1px solid #eef0ed; display:block;" />' +
          '</td>' +
          '<td style="padding:0 0 0 12px; vertical-align:middle; line-height:1.4;">' + displayName + '</td>' +
        '</tr></table>' +
      '</td>' +
      '<td style="padding:12px 10px; font-size:14px; color:#555555; text-align:center;">' + qty + '</td>' +
      '<td style="padding:12px 15px 12px 10px; font-size:14px; color:#888; text-align:right; font-weight:500; text-decoration:line-through;">' +
        (subtotal > 0 ? currency + " " + subtotal : "—") +
      '</td>' +
    '</tr>';
  }
  return html;
}

// 共用：已取消進度條 HTML
function buildCancelProgressBar(isTW) {
  var accentClr = isTW ? "#2980b9" : "#c0392b";
  var labels = isTW
    ? ["已收到訂單", "已付款", "已到台", "已完成取貨"]
    : ["已收到訂單", "已付款", "到港途中", "已寄出"];

  var nodes = "";
  for (var i = 0; i < 4; i++) {
    nodes += '<div style="position:relative; z-index:2; text-align:center; width:25%;">' +
      '<div style="width:11px; height:11px; background-color:#d1d5db; border-radius:50%; margin:0 auto; border:2px solid #ffffff;"></div>' +
      '<div style="font-size:11px; color:#888888; margin-top:8px;">' + labels[i] + '</div>' +
    '</div>';
  }

  return '<div style="margin:35px 0; padding:0 10px;">' +
    '<div style="position:relative; display:flex; justify-content:space-between; align-items:flex-start; width:100%;">' +
      '<div style="position:absolute; top:5px; left:0; right:0; height:1px; background-color:#e5e5e5; z-index:1;"></div>' +
      nodes +
    '</div>' +
    '<div style="margin-top:16px; text-align:center; padding:10px 0;">' +
      '<span style="display:inline-block; background:' + accentClr + '1a; border:1.5px solid ' + accentClr + '; color:' + accentClr + '; font-size:13px; font-weight:700; padding:6px 20px; border-radius:20px;">' +
        '✕ 訂單已取消' +
    '</span></div>' +
  '</div>';
}

// =================================================================
// sendCancelEmail：取消訂單通知（原因 A=未付款 / C=客人要求）
// =================================================================
function sendCancelEmail(ss, orderId, reason) {
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return { result: "error", message: "找不到訂單紀錄" };

  var rows = sheet.getDataRange().getValues();
  var targetRow = -1;
  for (var r = 1; r < rows.length; r++) {
    if (rows[r][0].toString() === orderId.toString()) { targetRow = r; break; }
  }
  if (targetRow === -1) return { result: "error", message: "找不到訂單編號" };

  var rowData    = rows[targetRow];
  var eventName  = rowData[2] ? rowData[2].toString().trim() : "代購活動";
  var summary    = rowData[3] ? rowData[3].toString() : "";
  var amount     = rowData[4] || 0;
  var orderDate  = rowData[1] ? new Date(rowData[1]) : new Date();
  var payMethod  = rowData[7] ? rowData[7].toString() : "";
  var custEmail  = rowData[10] ? rowData[10].toString().trim() : "";
  var pickupType = rowData[11] ? rowData[11].toString() : "";
  var isTW = pickupType.indexOf("賣貨便") !== -1 || payMethod.indexOf("郵局") !== -1;

  var d = orderDate;
  var fmtDate   = d.getFullYear() + "/" + (d.getMonth()+1) + "/" + d.getDate();
  var currency  = isTW ? "NT$" : "HK$";
  var accentClr = isTW ? "#2980b9" : "#c0392b";
  var bgColor   = isTW ? "#f7f6f3" : "#fbfbfa";

  var productSheet = getSheetByEventName(ss, eventName);
  var allItems = parseSummaryItems(summary);
  var itemsTableHtml = buildCancelItemsTable(allItems, productSheet, isTW);
  var progressHtml   = buildCancelProgressBar(isTW);

  var reasonTextHK = reason === "A"
    ? "由於我哋未能喺指定時間內收到付款，你嘅訂單已被取消。"
    : "根據你嘅要求，你嘅訂單已被取消。";
  var reasonTextTW = reason === "A"
    ? "由於我們未能於指定時間內收到匯款，您的訂單已被取消。"
    : "根據您的要求，您的訂單已被取消。";
  // A_deadline = 截單自動取消，不加「歡迎重新下單」
  var closingHK = (reason === "A_deadline") ? "" : "如仍有興趣，歡迎於收單期間內重新落單。";
  var closingTW = (reason === "A_deadline") ? "" : "如仍有興趣，歡迎於收單期間內重新下單。";
  var footerHK = "有任何問題歡迎直接回覆此 Email 或喺 Threads inbox 我哋查詢，謝謝！";
  var footerTW = "如有任何問題，歡迎直接回覆此 Email 或於 Threads 私訊我們查詢，謝謝！";

  var greeting  = isTW ? "您好！" : "你好～";
  var thanks    = isTW ? "感謝您於 886tw.81jp 訂購代購 <strong>" + eventName + "</strong>。"
                       : "感謝你早前喺 852hk.81jp 下單代購 <strong>" + eventName + "</strong>。";
  var reasonTxt = isTW ? reasonTextTW : reasonTextHK;
  var closing   = isTW ? closingTW : closingHK;
  var footer    = isTW ? footerTW : footerHK;

  var subject = "[訂單取消] 代購 " + eventName + " — " + (isTW ? "886tw.81jp" : "852hk.81jp");
  var divOpen = "<div style=\"font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width:600px; margin:0 auto; color:#333333; line-height:1.6; padding:20px; background-color:" + bgColor + ";\">";
  var htmlBody = divOpen +
    "<h2 style=\"font-size:18px; font-weight:500; border-bottom:1px solid #e5e5e5; padding-bottom:10px; color:#111111;\">代購 " + eventName + "</h2>" +
    "<p style=\"font-size:14px;\">" + greeting + "</p>" +
    "<p style=\"font-size:14px;\">" + thanks + "</p>" +
    "<p style=\"font-size:14px; background:#f2f2f2; padding:8px 12px; border-radius:4px; display:inline-block;\"><strong>訂單編號：</strong>" + orderId + "</p>" +
    progressHtml +
    "<p style=\"font-size:14px; color:#c62828; font-weight:500; margin-top:20px;\">" + reasonTxt + "</p>" +
    "<h3 style=\"font-size:14px; font-weight:bold; margin-top:20px; margin-bottom:10px; color:#444444;\">📋 訂單商品明細</h3>" +
    "<table style=\"width:100%; border-collapse:collapse; background:#ffffff; border:1px solid #eef0ed; border-radius:6px; overflow:hidden;\">" +
      "<thead><tr style=\"background-color:#fafbfa; border-bottom:1px solid #eef0ed;\">" +
        "<th style=\"padding:10px 15px; text-align:left; font-size:13px; color:#666666; font-weight:500;\">商品名稱與款式</th>" +
        "<th style=\"padding:10px; text-align:center; font-size:13px; color:#666666; font-weight:500; width:60px;\">數量</th>" +
        "<th style=\"padding:10px 15px; text-align:right; font-size:13px; color:#666666; font-weight:500; width:90px;\">金額</th>" +
      "</tr></thead>" +
      "<tbody>" + itemsTableHtml + "</tbody>" +
    "</table>" +
    "<p style=\"text-align:right; font-size:16px; font-weight:bold; margin-top:15px; color:#888; text-decoration:line-through;\">" + currency + " " + amount + "</p>" +
    "<p style=\"font-size:13px; color:#555; margin-top:20px;\">" + closing + "</p>" +
    "<div style=\"margin-top:30px; padding-top:15px; border-top:1px dashed #dddddd; font-size:12px; color:#777777; line-height:1.8;\">" +
      "<p style=\"margin:5px 0;\">" + footer + "</p>" +
    "</div>" +
    "<div style=\"margin-top:30px; font-size:13px; color:#333333; line-height:1.5;\">" +
      "<strong>" + (isTW ? "886tw.81jp" : "852hk.81jp") + "</strong><br>" +
      (isTW ? "<a href=\"https://www.threads.com/@886tw.81jp?igshid=NTc4MTIwNjQ2YQ==\" style=\"color:#2980b9;\">Threads @886tw.81jp</a> <a href=\"https://www.instagram.com/886tw.81jp?igsh=MW8zMmVncGVwNmd1dg%3D%3D&utm_source=qr\" style=\"color:#e1306c;\">Instagram @886tw.81jp</a>" : "<a href=\"https://www.threads.net/@852hk.81jp\" style=\"color:#111;\">Threads @852hk.81jp</a>") + "<br>" +
      "<div style=\"display:none;\">" + accentClr + "\">https://www.threads.net/@852hk.81jp</a>" +
    "</div>" +
    "</div>";

  sheet.getRange(targetRow + 1, 16).setValue("已取消");
  // 自動更新購貨紀錄：從 E（已付款）或 F（未付款）欄減去數量
  try {
    var cItems = parseSummaryItems(rowData[3] ? rowData[3].toString() : "");
    var cPaid  = (rowData[15] === "已收款" || rowData[15] === "已付款");
    updatePurchaseRecordQty(ss, eventName, cItems, cPaid ? 5 : 6, -1);
    // ── 即時更新 Supabase ordered_qty（退回庫存）──
    var cSummary = rowData[3] ? rowData[3].toString() : "";
    updateQtyOrderedInSupabase(stripEndPrefix(eventName), cSummary, -1, cPaid ? -1 : 0);
  } catch(ce) { Logger.log("取消訂單更新購貨紀錄失敗: " + ce.toString()); }

  if (custEmail && custEmail.indexOf("@") !== -1) {
    MailApp.sendEmail({ to: custEmail, subject: subject,
      body: "你好，訂單 " + orderId + " 已被取消。", htmlBody: htmlBody });
  }
  return { result: "success" };
}


// =================================================================
// sendHKShippingEmail：香港「到港途中」通知（貨品由日本寄出）
// =================================================================
function sendHKShippingEmail(ss, orderId, trackingNo) {
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return { result: "error", message: "找不到訂單紀錄" };
  var rows = sheet.getDataRange().getValues();
  var tr = -1;
  for (var r = 1; r < rows.length; r++) {
    if (rows[r][0].toString() === orderId.toString()) { tr = r; break; }
  }
  if (tr === -1) return { result: "error", message: "找不到訂單" };

  var rowData    = rows[tr];
  var eventName  = rowData[2] ? rowData[2].toString().trim() : "代購活動";
  var custEmail  = rowData[10] ? rowData[10].toString().trim() : "";
  var custName   = rowData[8] ? rowData[8].toString().trim() : "客人";
  var recipientName = rowData[25] ? rowData[25].toString().trim() : "-"; // Z欄
  var pickupCode    = rowData[12] ? rowData[12].toString().trim() : "-"; // M欄
  var pickupName    = rowData[13] ? rowData[13].toString().trim() : "-"; // N欄
  var pickupAddress = rowData[14] ? rowData[14].toString().trim() : "-"; // O欄
  var recipientPhone = rowData[9] ? rowData[9].toString().trim() : "-";  // J欄

  if (!custEmail || custEmail.indexOf("@") === -1) return { result: "error", message: "無效 Email" };

  var prog = "<div style='margin:28px 0;padding:0 10px;'><div style='position:relative;display:flex;justify-content:space-between;align-items:flex-start;'><div style='position:absolute;top:5px;left:0;right:0;height:1px;background:#e5e5e5;z-index:1;'></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>已收到訂單</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>已付款</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>到港途中</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#d1d5db;border-radius:50%;margin:0 auto;border:2px solid #fff;'></div><div style='font-size:11px;color:#888;margin-top:7px;'>已寄出</div></div>" +
    "</div></div>";

  var subject = "[到港途中] " + eventName + " - 852hk.81jp";
  var body = "<div style='font-family:Helvetica Neue,Arial,sans-serif;max-width:600px;margin:0 auto;padding:20px;background:#fbfbfa;color:#333;line-height:1.6;'>" +
    "<h2 style='font-size:18px;font-weight:500;border-bottom:1px solid #e5e5e5;padding-bottom:10px;color:#111;'>代購 " + eventName + "</h2>" +
    "<p style='font-size:14px;'>你好，" + custName + "！</p>" +
    "<p style='font-size:14px;color:#111;font-weight:500;margin-top:8px;'>📦 你嘅代購商品已由日本寄出，正在運送往香港途中！</p>" +
    "<p style='font-size:14px;background:#f2f2f2;padding:8px 12px;border-radius:4px;display:inline-block;margin-top:10px;'><strong>訂單編號：</strong>" + orderId + "</p>" +
    prog +
    "<div style='background:#f0f7ff;border-left:4px solid #111;padding:14px;margin:16px 0;border-radius:4px;'>" +
      "<div style='font-weight:700;font-size:14px;margin-bottom:8px;'>📮 運單號碼</div>" +
      "<div style='font-family:monospace;font-size:18px;font-weight:700;color:#111;background:#fff;border:2px dashed #111;display:inline-block;padding:6px 16px;border-radius:6px;letter-spacing:3px;'>" + (trackingNo || "—") + "</div>" +
    "</div>" +
    "<h3 style='font-size:14px;font-weight:700;margin:20px 0 8px;color:#444;'>📍 請確認易寄取收件資料</h3>" +
    "<div style='background:#fff;border:1px solid #eef0ed;padding:14px;border-radius:6px;font-size:13px;color:#555;line-height:1.9;'>" +
      "<div><strong>收件人：</strong>" + recipientName + "</div>" +
      "<div><strong>聯絡電話：</strong>" + recipientPhone + "</div>" +
      "<div><strong>自提點編號：</strong>" + pickupCode + "</div>" +
      "<div><strong>自提點名稱：</strong>" + pickupName + "</div>" +
      "<div><strong>自提點地址：</strong>" + pickupAddress + "</div>" +
    "</div>" +
    "<div style='background:#fff8e1;border-left:4px solid #f59e0b;padding:12px 14px;border-radius:4px;margin:14px 0;font-size:13px;color:#78350f;line-height:1.7;'>" +
      "⚠️ 如以上地址有誤，請於收到此 Email 後 <strong>24小時內</strong> 盡快回覆，以便我哋安排更改。<br>" +
      "如地址正確，無需回覆，我哋將以上述地址寄出。" +
    "</div>" +
    "<div style='margin-top:28px;padding-top:12px;border-top:1px dashed #ddd;font-size:12px;color:#777;line-height:1.8;'>" +
      "<p>如有任何問題，歡迎直接回覆此 Email 或到 Threads inbox 我哋查詢，謝謝！</p>" + "<div style='margin-top:14px;padding:12px;background:#f5f5f5;border-radius:8px;text-align:center;'><a href='https://www.threads.net/@852hk.81jp' style='display:inline-block;background:#111;color:#fff;font-size:12px;font-weight:700;padding:6px 14px;border-radius:14px;text-decoration:none;'>Threads @852hk.81jp</a></div>" + "<p style='text-align:center;margin-top:10px;'><strong>852hk.81jp</strong></p>" +
    "</div></div>";

  MailApp.sendEmail({ to: custEmail, subject: subject, body: "你好，你嘅代購商品正在運送途中，訂單編號：" + orderId, htmlBody: body });
  return { result: "success" };
}

// =================================================================
// sendHKDeliveredEmail：香港「已寄出」通知（提供易寄取運單號）
// =================================================================
function sendHKDeliveredEmail(ss, orderId, trackingNo) {
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return { result: "error", message: "找不到訂單紀錄" };
  var rows = sheet.getDataRange().getValues();
  var tr = -1;
  for (var r = 1; r < rows.length; r++) {
    if (rows[r][0].toString() === orderId.toString()) { tr = r; break; }
  }
  if (tr === -1) return { result: "error", message: "找不到訂單" };

  var rowData    = rows[tr];
  var eventName  = rowData[2] ? rowData[2].toString().trim() : "代購活動";
  var custEmail  = rowData[10] ? rowData[10].toString().trim() : "";
  var custName   = rowData[8] ? rowData[8].toString().trim() : "客人";
  var pickupName = rowData[13] ? rowData[13].toString().trim() : "-";

  if (!custEmail || custEmail.indexOf("@") === -1) return { result: "error", message: "無效 Email" };

  var prog = "<div style='margin:28px 0;padding:0 10px;'><div style='position:relative;display:flex;justify-content:space-between;align-items:flex-start;'><div style='position:absolute;top:5px;left:0;right:0;height:1px;background:#e5e5e5;z-index:1;'></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>已收到訂單</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>已付款</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>到港途中</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>已寄出</div></div>" +
    "</div></div>";

  var subject = "🎉 [已寄出] " + eventName + " - 852hk.81jp";
  var body = "<div style='font-family:Helvetica Neue,Arial,sans-serif;max-width:600px;margin:0 auto;padding:20px;background:#fbfbfa;color:#333;line-height:1.6;'>" +
    "<h2 style='font-size:18px;font-weight:500;border-bottom:1px solid #e5e5e5;padding-bottom:10px;color:#111;'>代購 " + eventName + "</h2>" +
    "<p style='font-size:14px;'>你好，" + custName + "！</p>" +
    "<p style='font-size:14px;color:#111;font-weight:500;margin-top:8px;'>🎉 你嘅代購商品已寄出，請憑以下運單號碼到 <strong>" + pickupName + "</strong> 取件！</p>" +
    "<p style='font-size:14px;background:#f2f2f2;padding:8px 12px;border-radius:4px;display:inline-block;margin-top:10px;'><strong>訂單編號：</strong>" + orderId + "</p>" +
    prog +
    "<div style='background:#f0f7ff;border-left:4px solid #111;padding:14px;margin:16px 0;border-radius:4px;'>" +
      "<div style='font-weight:700;font-size:14px;margin-bottom:8px;'>📮 易寄取運單號碼</div>" +
      "<div style='font-family:monospace;font-size:22px;font-weight:700;color:#111;background:#fff;border:2px dashed #111;display:inline-block;padding:8px 20px;border-radius:6px;letter-spacing:4px;'>" + (trackingNo || "—") + "</div>" +
      "<p style='font-size:12px;color:#888;margin-top:8px;'>請將此號碼出示予自提點人員，或於易寄取 App 內查詢。</p>" +
    "</div>" +
    "<div style='background:#f0fff4;border-left:4px solid #10b981;padding:12px 14px;border-radius:4px;margin:14px 0;font-size:13px;color:#065f46;line-height:1.7;'>" +
      "✅ 請儘快前往 <strong>" + pickupName + "</strong> 取件，避免超過保管期限。" +
    "</div>" +
    "<div style='margin-top:28px;padding-top:12px;border-top:1px dashed #ddd;font-size:12px;color:#777;line-height:1.8;'>" +
      "<p>如有任何問題，歡迎直接回覆此 Email 或到 Threads inbox 我哋查詢，謝謝！</p>" + "<div style='margin-top:14px;padding:12px;background:#f5f5f5;border-radius:8px;text-align:center;'><a href='https://www.threads.net/@852hk.81jp' style='display:inline-block;background:#111;color:#fff;font-size:12px;font-weight:700;padding:6px 14px;border-radius:14px;text-decoration:none;'>Threads @852hk.81jp</a></div>" + "<p style='text-align:center;margin-top:10px;'><strong>852hk.81jp</strong></p>" +
    "</div></div>";

  MailApp.sendEmail({ to: custEmail, subject: subject, body: "你好，你嘅代購商品已寄出，訂單編號：" + orderId + "，運單號碼：" + trackingNo, htmlBody: body });
  // 更新狀態為已完結
  sheet.getRange(tr + 1, 16).setValue("已完結");
  return { result: "success" };
}

// =================================================================
// sendStockoutEmail：缺貨取消通知（部分或全部，自動重算金額）
// =================================================================
function sendStockoutEmail(ss, orderId, outOfStockItems) {
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return { result: "error", message: "找不到訂單紀錄" };

  var rows = sheet.getDataRange().getValues();
  var targetRow = -1;
  for (var r = 1; r < rows.length; r++) {
    if (rows[r][0].toString() === orderId.toString()) { targetRow = r; break; }
  }
  if (targetRow === -1) return { result: "error", message: "找不到訂單編號" };

  var rowData    = rows[targetRow];
  var eventName  = rowData[2] ? rowData[2].toString().trim() : "代購活動";
  var summary    = rowData[3] ? rowData[3].toString() : "";
  var orderDate  = rowData[1] ? new Date(rowData[1]) : new Date();
  var payMethod  = rowData[7] ? rowData[7].toString() : "";
  var custEmail  = rowData[10] ? rowData[10].toString().trim() : "";
  var pickupType = rowData[11] ? rowData[11].toString() : "";
  var isTW = pickupType.indexOf("賣貨便") !== -1 || payMethod.indexOf("郵局") !== -1;

  var d = orderDate;
  var fmtDate   = d.getFullYear() + "/" + (d.getMonth()+1) + "/" + d.getDate();
  var currency  = isTW ? "NT$" : "HK$";
  var priceCol  = isTW ? 8 : 4;
  var accentClr = isTW ? "#2980b9" : "#c0392b";
  var bgColor   = isTW ? "#f7f6f3" : "#fbfbfa";

  var productSheet = getSheetByEventName(ss, eventName);

  // 讀取商品價格
  var priceMap = {};
  if (productSheet) {
    var pRows = productSheet.getDataRange().getValues();
    for (var p = 1; p < pRows.length; p++) {
      var pn = pRows[p][0] ? pRows[p][0].toString().trim() : "";
      if (pn) priceMap[pn] = parseFloat(pRows[p][priceCol]) || 0;
    }
  }

  var allItems = parseSummaryItems(summary);

  // 區分缺貨和保留商品
  var outKeys = {};
  for (var oi = 0; oi < outOfStockItems.length; oi++) {
    outKeys[outOfStockItems[oi].name + "|" + outOfStockItems[oi].sub] = true;
  }
  var remainItems = [], removedItems = [];
  for (var ai = 0; ai < allItems.length; ai++) {
    var key = allItems[ai].name + "|" + allItems[ai].sub;
    if (outKeys[key]) removedItems.push(allItems[ai]);
    else remainItems.push(allItems[ai]);
  }
  var isAllOut = remainItems.length === 0;

  // 重新計算金額
  function calcTotal(items) {
    var total = 0;
    for (var i = 0; i < items.length; i++) {
      total += (priceMap[items[i].name] || 0) * items[i].qty;
    }
    if (!isTW) total += (total >= 500 ? 0 : 10);
    return total;
  }
  var newAmount = isAllOut ? 0 : calcTotal(remainItems);

  var progressHtml = buildCancelProgressBar(isTW);

  // 建立缺貨商品表格（有刪除線）
  var removedTableHtml = buildCancelItemsTable(removedItems, productSheet, isTW);

  // 建立保留商品表格（正常顯示）
  var remainTableHtml = "";
  if (!isAllOut && productSheet) {
    var fallbackImg = "data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%221%22%20height%3D%221%22%3E%3Crect%20width%3D%221%22%20height%3D%221%22%20fill%3D%22%23111%22%2F%3E%3C%2Fsvg%3E";
    var photoMap = {};
    var pRows2 = productSheet.getDataRange().getValues();
    for (var pp = 1; pp < pRows2.length; pp++) {
      var ppn = pRows2[pp][0] ? pRows2[pp][0].toString().trim() : "";
      if (ppn) photoMap[ppn] = pRows2[pp][5] ? pRows2[pp][5].toString().trim() : "";
    }
    for (var ri = 0; ri < remainItems.length; ri++) {
      var item = remainItems[ri];
      var dn = item.name + (item.sub ? " (" + item.sub + ")" : "");
      var qty = item.qty || 1;
      var photo = photoMap[item.name] || "";
      var imgUrl = (photo && photo.indexOf("http") === 0) ? photo : fallbackImg;
      var price = priceMap[item.name] || 0;
      var subtotal = price * qty;
      remainTableHtml +=
        '<tr style="border-bottom:1px solid #eeeeee;">' +
          '<td style="padding:12px 10px 12px 15px; font-size:14px; color:#333333;">' +
            '<table style="width:100%; border-collapse:collapse; border:none;"><tr>' +
              '<td style="padding:0; width:50px; vertical-align:middle;">' +
                '<img src="' + imgUrl + '" style="width:50px; height:50px; object-fit:cover; border-radius:4px; border:1px solid #eef0ed; display:block;" />' +
              '</td>' +
              '<td style="padding:0 0 0 12px; vertical-align:middle; line-height:1.4;">' + dn + '</td>' +
            '</tr></table>' +
          '</td>' +
          '<td style="padding:12px 10px; font-size:14px; color:#555555; text-align:center;">' + qty + '</td>' +
          '<td style="padding:12px 15px 12px 10px; font-size:14px; color:' + accentClr + '; text-align:right; font-weight:500;">' +
            (subtotal > 0 ? currency + " " + subtotal : "—") +
          '</td>' +
        '</tr>';
    }
  }

  var greeting = isTW ? "您好！" : "你好～";
  var thanks   = isTW ? "感謝您於 886tw.81jp 訂購代購 <strong>" + eventName + "</strong>。"
                      : "感謝你早前喺 852hk.81jp 下單代購 <strong>" + eventName + "</strong>。";
  var footer   = isTW ? "如有任何問題，歡迎直接回覆此 Email 或於 Threads 私訊我們查詢，謝謝！"
                      : "有任何問題歡迎直接回覆此 Email 或喺 Threads inbox 我哋查詢，謝謝！";
  var closing  = isTW ? "如有其他感興趣的商品，歡迎於收單期間內重新下單。"
                      : "如有其他心水商品，歡迎於收單期間內重新落單。";

  var mainText = isAllOut
    ? (isTW ? "非常抱歉，由於您所訂購的商品已全部缺貨，您的訂單已被取消。"
            : "非常抱歉，由於你所訂購的商品已全部缺貨，你嘅訂單已被取消。")
    : (isTW ? "非常抱歉，由於以下商品已缺貨，相關部分已從您的訂單中移除，訂單金額已更新。"
            : "非常抱歉，由於以下商品已缺貨，相關部分已從你嘅訂單中移除，訂單金額已更新。");

  var subject = "[訂單取消] 代購 " + eventName + " — 852hk.81jp";

  var remainSection = "";
  if (!isAllOut) {
    remainSection =
      "<h3 style=\"font-size:14px; font-weight:bold; margin-top:20px; margin-bottom:10px; color:#444444;\">✅ 保留商品</h3>" +
      "<table style=\"width:100%; border-collapse:collapse; background:#ffffff; border:1px solid #eef0ed; border-radius:6px; overflow:hidden;\">" +
        "<thead><tr style=\"background-color:#fafbfa; border-bottom:1px solid #eef0ed;\">" +
          "<th style=\"padding:10px 15px; text-align:left; font-size:13px; color:#666666; font-weight:500;\">商品名稱與款式</th>" +
          "<th style=\"padding:10px; text-align:center; font-size:13px; color:#666666; font-weight:500; width:60px;\">數量</th>" +
          "<th style=\"padding:10px 15px; text-align:right; font-size:13px; color:#666666; font-weight:500; width:90px;\">小計</th>" +
        "</tr></thead>" +
        "<tbody>" + remainTableHtml + "</tbody>" +
      "</table>" +
      "<p style=\"text-align:right; font-size:16px; font-weight:bold; margin-top:12px; color:#111111;\">" +
        "更新後金額：<span style=\"font-size:20px; color:" + accentClr + ";\">" + currency + " " + newAmount + "</span>" +
      "</p>";
  }

  var divOpen2 = "<div style=\"font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width:600px; margin:0 auto; color:#333333; line-height:1.6; padding:20px; background-color:" + bgColor + ";\">";
  var htmlBody = divOpen2 +
    "<h2 style=\"font-size:18px; font-weight:500; border-bottom:1px solid #e5e5e5; padding-bottom:10px; color:#111111;\">代購 " + eventName + "</h2>" +
    "<p style=\"font-size:14px;\">" + greeting + "</p>" +
    "<p style=\"font-size:14px;\">" + thanks + "</p>" +
    "<p style=\"font-size:14px; background:#f2f2f2; padding:8px 12px; border-radius:4px; display:inline-block;\"><strong>訂單編號：</strong>" + orderId + "</p>" +
    progressHtml +
    "<p style=\"font-size:14px; color:#c62828; font-weight:500; margin-top:20px;\">" + mainText + "</p>" +
    "<h3 style=\"font-size:14px; font-weight:bold; margin-top:20px; margin-bottom:10px; color:#444444;\">❌ 缺貨商品</h3>" +
    "<table style=\"width:100%; border-collapse:collapse; background:#fff5f5; border:1px solid #fca5a5; border-radius:6px; overflow:hidden;\">" +
      "<thead><tr style=\"background-color:#fef2f2; border-bottom:1px solid #fca5a5;\">" +
        "<th style=\"padding:10px 15px; text-align:left; font-size:13px; color:#666666; font-weight:500;\">商品名稱與款式</th>" +
        "<th style=\"padding:10px; text-align:center; font-size:13px; color:#666666; font-weight:500; width:60px;\">數量</th>" +
        "<th style=\"padding:10px 15px; text-align:right; font-size:13px; color:#666666; font-weight:500; width:90px;\">金額</th>" +
      "</tr></thead>" +
      "<tbody>" + removedTableHtml + "</tbody>" +
    "</table>" +
    remainSection +
    "<p style=\"font-size:13px; color:#555; margin-top:20px;\">" + closing + "</p>" +
    "<div style=\"margin-top:30px; padding-top:15px; border-top:1px dashed #dddddd; font-size:12px; color:#777777;\">" +
      "<p style=\"margin:5px 0;\">" + footer + "</p>" +
    "</div>" +
    "<div style=\"margin-top:30px; font-size:13px; color:#333333; line-height:1.5;\">" +
      "<strong>852hk.81jp</strong><br>" +
      "Threads: <a href=\"https://www.threads.net/@852hk.81jp\" target=\"_blank\" style=\"color:" + accentClr + "\">https://www.threads.net/@852hk.81jp</a>" +
    "</div>" +
    "</div>";

  sheet.getRange(targetRow + 1, 5).setValue(newAmount);
  sheet.getRange(targetRow + 1, 16).setValue("已取消");

  if (custEmail && custEmail.indexOf("@") !== -1) {
    MailApp.sendEmail({ to: custEmail, subject: subject,
      body: "關於訂單 " + orderId + " 的缺貨通知，詳情請查看 HTML 郵件。", htmlBody: htmlBody });
  }
  return { result: "success", newAmount: newAmount, isAllOut: isAllOut };
}

// =================================================================
// getProfitData：盈利頁面所需的完整資料
// =================================================================
function getProfitData(ss, eventName) {
  var purchaseSheet = getPurchaseSheet(ss);
  var profitSheet   = getProfitSheet(ss);
  var orderSheet    = ss.getSheetByName("訂單紀錄");
  var cleanEventName = stripEndPrefix(eventName);
  var eventSheet    = getSheetByEventName(ss, eventName);

  // ── 讀取盈利紀錄（成本 + 匯率）──
  var totalCost = 0, twdRate = 3.97, settleDate = "";
  if (profitSheet) {
    var pRows = profitSheet.getDataRange().getValues();
    for (var pr = 1; pr < pRows.length; pr++) {
      var prName = pRows[pr][0] ? pRows[pr][0].toString().trim() : "";
      if (prName === cleanEventName || prName === eventName) {
        totalCost  = parseFloat(pRows[pr][1]) || 0;
        twdRate    = parseFloat(pRows[pr][2]) || 3.97;
        settleDate = pRows[pr][3] ? pRows[pr][3].toString() : "";
        break;
      }
    }
  }

  // ── 讀取商品定價（HKD / TWD / 圖片）──
  var priceMap = {}; // { "商品名": { hkd, twd, photo, yen } }
  if (eventSheet) {
    var eRows = eventSheet.getDataRange().getValues();
    for (var er = 1; er < eRows.length; er++) {
      var en = eRows[er][0] ? eRows[er][0].toString().trim() : "";
      if (!en) continue;
      priceMap[en] = {
        yen:   parseFloat(eRows[er][2]) || 0,
        hkd:   parseFloat(eRows[er][4]) || 0,
        twd:   parseFloat(eRows[er][8]) || 0,
        photo: eRows[er][5] ? eRows[er][5].toString().trim() : ""
      };
    }
  }

  // ── 讀取購貨紀錄（D=日幣, G=已購買, H=已結算）──
  var settledMap   = {}; // { "商品名|款式": 已結算qty }
  var yenFromSheet = {}; // { "商品名": 日幣單價 } from D欄 / C欄(per-event)
  // 優先讀 per-event sheet
  var perEvProfitSheet = getEventPurchaseSheet(ss, eventName, false);
  if (perEvProfitSheet) {
    // per-event：A=名稱, B=款式, C=日幣, D=已付款, E=未付款, F=已購買, G=已結算
    var purRows = perEvProfitSheet.getDataRange().getValues();
    for (var pur = 1; pur < purRows.length; pur++) {
      if (!purRows[pur][0]) continue;
      var purName    = purRows[pur][0] ? purRows[pur][0].toString().trim() : "";
      var purSub     = purRows[pur][1] ? purRows[pur][1].toString().trim() : "";
      var purYen     = Number(purRows[pur][2]) || 0; // C欄(index2)=日幣
      var purSettled = Number(purRows[pur][6]) || 0; // G欄(index6)=已結算
      if (purYen > 0)     yenFromSheet[purName] = purYen;
      if (purSettled > 0) settledMap[purName + "|" + purSub] = purSettled;
    }
  } else if (purchaseSheet) {
    // 舊版總表：A=Event, B=名稱, C=款式, D=日幣, E=已付款, F=未付款, G=已購買, H=已結算
    var purRows2 = purchaseSheet.getDataRange().getValues();
    for (var pur2 = 1; pur2 < purRows2.length; pur2++) {
      if (!purRows2[pur2][0]) continue;
      var purEvent = purRows2[pur2][0].toString().trim();
      if (purEvent !== cleanEventName) continue;
      var purName2   = purRows2[pur2][1] ? purRows2[pur2][1].toString().trim() : "";
      var purSub2    = purRows2[pur2][2] ? purRows2[pur2][2].toString().trim() : "";
      var purYen2    = Number(purRows2[pur2][3]) || 0;
      var purSettled2= Number(purRows2[pur2][7]) || 0;
      if (purYen2 > 0)     yenFromSheet[purName2] = purYen2;
      if (purSettled2 > 0) settledMap[purName2 + "|" + purSub2] = purSettled2;
    }
  }

  // ── 讀取已付款訂單，統計 HK / TW 各款式數量 ──
  var hkQtyMap = {}, twQtyMap = {}; // { "商品名|款式": qty }
  if (orderSheet) {
    var oRowsP = orderSheet.getDataRange().getValues();
    for (var oi = 1; oi < oRowsP.length; oi++) {
      var orEvent  = oRowsP[oi][2] ? oRowsP[oi][2].toString().trim() : "";
      var orStatus = oRowsP[oi][15] ? oRowsP[oi][15].toString().trim() : "";
      var orPay    = oRowsP[oi][7] ? oRowsP[oi][7].toString() : "";
      var orPickup = oRowsP[oi][11] ? oRowsP[oi][11].toString() : "";
      if (orEvent !== cleanEventName) continue;
      if (orStatus !== "已收款" && orStatus !== "已付款") continue;
      var isTW = orPickup.indexOf("賣貨便") !== -1 || orPay.indexOf("郵局") !== -1;
      var pItems = parseSummaryItems(oRowsP[oi][3] ? oRowsP[oi][3].toString() : "");
      for (var ii = 0; ii < pItems.length; ii++) {
        var iKey = pItems[ii].name + "|" + pItems[ii].sub;
        if (isTW) twQtyMap[iKey] = (twQtyMap[iKey] || 0) + pItems[ii].qty;
        else      hkQtyMap[iKey] = (hkQtyMap[iKey] || 0) + pItems[ii].qty;
      }
    }
  }

  // ── 計算各商品比例成本 ──
  var totalSettledYen = 0;
  for (var key in settledMap) {
    var sName = key.split("|")[0];
    // 優先使用購貨紀錄 D欄日幣
    var sYen = yenFromSheet[sName] || (priceMap[sName] ? priceMap[sName].yen : 0);
    totalSettledYen += sYen * settledMap[key];
  }

  // ── 組合回傳資料 ──
  var items = [];
  for (var key in settledMap) {
    var parts   = key.split("|");
    var name    = parts[0];
    var sub     = parts[1] || "";
    var qty     = settledMap[key];
    var price   = priceMap[name] || { yen: 0, hkd: 0, twd: 0, photo: "" };
    // 優先使用購貨紀錄 D欄的日幣（最準確），fallback 到 event sheet
    var yenEach  = yenFromSheet[name] || price.yen || 0;
    var yenTotal  = yenEach * qty;
    var costProp  = totalSettledYen > 0 ? totalCost * yenTotal / totalSettledYen : 0;
    var hkQty   = hkQtyMap[key] || 0;
    var twQty   = twQtyMap[key] || 0;
    var hkRev   = price.hkd * hkQty;
    var twRev   = twdRate > 0 ? (price.twd / twdRate) * twQty : 0;
    var profit  = hkRev + twRev - costProp;
    items.push({
      name: name, sub: sub, qty: qty,
      photo: price.photo,
      yenEach: price.yen, yenTotal: yenTotal,
      hkdPrice: price.hkd, twdPrice: price.twd,
      hkQty: hkQty, twQty: twQty,
      hkRev: Math.round(hkRev * 100) / 100,
      twRev: Math.round(twRev * 100) / 100,
      costProp: Math.round(costProp * 100) / 100,
      profit: Math.round(profit * 100) / 100
    });
  }

  return {
    event: eventName, totalCost: totalCost, twdRate: twdRate,
    settleDate: settleDate, totalSettledYen: totalSettledYen,
    items: items
  };
}

// =================================================================
// saveProfitCost：儲存港幣總成本到盈利紀錄
// =================================================================
function saveProfitCost(ss, eventName, totalCost, twdRate) {
  var cleanEventName = stripEndPrefix(eventName);
  var sheet = getProfitSheet(ss);
  if (!sheet) {
    sheet = ss.insertSheet("盈利紀錄");
    sheet.appendRow(["Event名稱", "港幣總成本", "TWD匯率", "結算日期"]);
  }
  var rows = sheet.getDataRange().getValues();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm");
  for (var r = 1; r < rows.length; r++) {
    var storedName = rows[r][0] ? rows[r][0].toString().trim() : "";
    if (storedName === cleanEventName || storedName === eventName) {
      sheet.getRange(r + 1, 1).setValue(cleanEventName); // 順便修正為 clean name
      sheet.getRange(r + 1, 2).setValue(totalCost);
      sheet.getRange(r + 1, 3).setValue(twdRate);
      sheet.getRange(r + 1, 4).setValue(today);
      return;
    }
  }
  sheet.appendRow([cleanEventName, totalCost, twdRate, today]);
}

// =================================================================
// updatePurchaseRecord：更新購貨紀錄的 E(已付款) 或 F(未付款) 欄
// delta = 增減數量, colIndex = 5(E=已付款) 或 6(F=未付款)
// =================================================================
function updatePurchaseRecordQty(ss, eventName, items, colIndex, delta) {
  var cleanEventName = stripEndPrefix(eventName);
  // colIndex: 5=已付款, 6=未付款（舊版總表欄位）
  // 優先使用 per-event sheet（[Data](event)購貨紀錄）
  var perEventSheet = getEventPurchaseSheet(ss, eventName, false);
  var usePerEvent = !!perEventSheet;
  var sheet = usePerEvent ? perEventSheet : getPurchaseSheet(ss);
  if (!sheet) return;

  var rows = sheet.getDataRange().getValues();

  // 從 event sheet 取日幣
  var yenMap = {};
  var eventSheet = getSheetByEventName(ss, eventName);
  if (eventSheet) {
    var eRows = eventSheet.getDataRange().getValues();
    for (var er = 1; er < eRows.length; er++) {
      var en = eRows[er][0] ? eRows[er][0].toString().trim() : "";
      if (en) yenMap[en] = parseFloat(eRows[er][2]) || 0;
    }
  }

  // per-event 欄位：D=4(已付款), E=5(未付款)；舊版：E=5, F=6
  // colIndex 傳入為舊版欄號(5或6)，per-event 需減1
  var actualCol = usePerEvent ? (colIndex - 1) : colIndex;
  var nameIdx   = usePerEvent ? 0 : 1;
  var subIdx    = usePerEvent ? 1 : 2;

  for (var k = 0; k < items.length; k++) {
    var item = items[k];
    var itemSub  = item.sub  ? item.sub.toString().trim()  : "";
    var itemName = item.name ? item.name.toString().trim() : "";
    if (!itemName) continue;

    var found = false;
    for (var r = 1; r < rows.length; r++) {
      if (!rows[r][nameIdx]) continue;
      if (!usePerEvent) {
        var rEv = rows[r][0] ? rows[r][0].toString().trim() : "";
        if (rEv !== cleanEventName) continue;
      }
      var rNm  = rows[r][nameIdx] ? rows[r][nameIdx].toString().trim() : "";
      var rSb  = rows[r][subIdx]  ? rows[r][subIdx].toString().trim()  : "";
      if (rNm === itemName && rSb === itemSub) {
        var cur = Number(rows[r][actualCol - 1]) || 0;
        var newVal = Math.max(0, cur + delta * item.qty);
        sheet.getRange(r + 1, actualCol).setValue(newVal);
        rows[r][actualCol - 1] = newVal;
        found = true;
        break;
      }
    }
    if (!found && delta > 0) {
      var yenEach = yenMap[itemName] || 0;
      if (yenEach <= 0) {
        Logger.log("updatePurchaseRecordQty: 找不到商品日幣，跳過: " + itemName);
        continue;
      }
      if (usePerEvent) {
        var nr = [itemName, itemSub, yenEach, 0, 0, 0, 0];
        nr[actualCol - 1] = delta * item.qty;
        sheet.appendRow(nr);
        rows.push(nr);
      } else {
        var nr2 = [eventName, itemName, itemSub, yenEach, 0, 0, 0, 0];
        nr2[actualCol - 1] = delta * item.qty;
        sheet.appendRow(nr2);
        rows.push(nr2);
      }
    }
  }
}

// =================================================================
// initSyncPurchaseRecord：從訂單紀錄重新同步購貨紀錄 E/F 欄
// 用於遷移現有資料，以及每次重建時使用
// =================================================================
function initSyncPurchaseRecord(ss, eventName) {
  var cleanEventName = stripEndPrefix(eventName);
  // 優先使用 per-event sheet，fallback 到舊版總表
  var perEvSheet = getEventPurchaseSheet(ss, eventName, false);
  var usePerEvent = !!perEvSheet;
  var sheet = usePerEvent ? perEvSheet : getPurchaseSheet(ss);
  var orderSheet = ss.getSheetByName("訂單紀錄");
  if (!sheet || !orderSheet) return;

  // 從訂單紀錄統計 paidQty 和 unpaidQty
  var paidMap = {}, unpaidMap = {};
  var oRows = orderSheet.getDataRange().getValues();
  for (var r = 1; r < oRows.length; r++) {
    if (!oRows[r][0]) continue;
    var rowEvent  = oRows[r][2] ? oRows[r][2].toString().trim() : "";
    var rowStatus = oRows[r][15] ? oRows[r][15].toString().trim() : "";
    if (rowEvent !== cleanEventName) continue;
    if (rowStatus === "已取消") continue;
    var isPaid = (rowStatus === "已收款" || rowStatus === "已付款");
    var parsed = parseSummaryItems(oRows[r][3] ? oRows[r][3].toString() : "");
    for (var k = 0; k < parsed.length; k++) {
      var key = parsed[k].name + "|" + parsed[k].sub;
      if (isPaid) paidMap[key] = (paidMap[key] || 0) + parsed[k].qty;
      else unpaidMap[key] = (unpaidMap[key] || 0) + parsed[k].qty;
    }
  }

  // 從 event sheet 取日幣單價
  var yenMap = {};
  var eventSheet = getSheetByEventName(ss, eventName);
  if (eventSheet) {
    var eRows = eventSheet.getDataRange().getValues();
    for (var er = 1; er < eRows.length; er++) {
      var en = eRows[er][0] ? eRows[er][0].toString().trim() : "";
      if (en) yenMap[en] = parseFloat(eRows[er][2]) || 0;
    }
  }

  // 欄位索引（per-event：A=名稱, B=款式, C=日幣, D=已付款, E=未付款）
  // 舊版：A=Event, B=名稱, C=款式, D=日幣, E=已付款, F=未付款
  var nameIdx = usePerEvent ? 0 : 1;
  var subIdx  = usePerEvent ? 1 : 2;
  var yenCol  = usePerEvent ? 3 : 4; // 1-indexed
  var paidCol = usePerEvent ? 4 : 5;
  var unpaidCol = usePerEvent ? 5 : 6;

  var rows = sheet.getDataRange().getValues();
  var allKeys = {};
  for (var k in paidMap) allKeys[k] = true;
  for (var k in unpaidMap) allKeys[k] = true;

  for (var key in allKeys) {
    var parts = key.split("|");
    var name = parts[0], sub = parts[1] || "";
    var paid   = paidMap[key]   || 0;
    var unpaid = unpaidMap[key] || 0;
    var yen    = yenMap[name]   || 0;
    var found  = false;

    for (var r = 1; r < rows.length; r++) {
      if (!rows[r][nameIdx]) continue;
      if (!usePerEvent) {
        var rEv = rows[r][0] ? rows[r][0].toString().trim() : "";
        if (rEv !== cleanEventName) continue;
      }
      var rName = rows[r][nameIdx].toString().trim();
      var rSub  = rows[r][subIdx] ? rows[r][subIdx].toString().trim() : "";
      var cleanSub = sub ? sub.toString().trim() : "";
      if (rName === name && rSub === cleanSub) {
        if (yen > 0) sheet.getRange(r + 1, yenCol).setValue(yen);
        sheet.getRange(r + 1, paidCol).setValue(paid);
        sheet.getRange(r + 1, unpaidCol).setValue(unpaid);
        found = true;
        break;
      }
    }
    if (!found) {
      if (yen <= 0) {
        Logger.log("initSyncPurchaseRecord: 找不到商品日幣，跳過: " + name);
        continue;
      }
      var newRow;
      if (usePerEvent) {
        newRow = [name, sub, yen, paid, unpaid, 0, 0];
      } else {
        newRow = [eventName, name, sub, yen, paid, unpaid, 0, 0];
      }
      sheet.appendRow(newRow);
      rows.push(newRow);
    }
  }

  // 同時確保有資料的行補齊日幣單價
  rows = sheet.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    if (!rows[r][nameIdx]) continue;
    if (!usePerEvent) {
      var rEv2 = rows[r][0] ? rows[r][0].toString().trim() : "";
      if (rEv2 !== cleanEventName) continue;
    }
    var rNm = rows[r][nameIdx].toString().trim();
    if (!rows[r][yenCol - 1] && yenMap[rNm]) {
      sheet.getRange(r + 1, yenCol).setValue(yenMap[rNm]);
    }
  }
}

// =================================================================
// sendShoppingList：發送「未購買」購物清單 email 給管理員自己
// =================================================================
function sendShoppingList(ss, eventName, items) {
  // items: [{ name, sub, qty, yenEach }] - 不傳 photo，從 event sheet 重新讀取
  // getEffectiveUser() 在 Execute as Me 部署下可靠地返回 owner email
  var adminEmail = Session.getEffectiveUser().getEmail();
  if (!adminEmail) {
    // fallback：用 owner email
    adminEmail = PropertiesService.getScriptProperties().getProperty("ADMIN_EMAIL") || "";
    if (!adminEmail) {
      Logger.log("sendShoppingList: 無法取得 adminEmail");
      return { result: "error", message: "無法取得管理員 email" };
    }
  }
  Logger.log("sendShoppingList: 發送至 " + adminEmail + "，共 " + items.length + " 款");

  // 從 event sheet 讀取商品圖片（避免前端傳 base64 造成 body 過大）
  var photoMap = {};
  var eventSheet = getSheetByEventName(ss, eventName);
  if (eventSheet) {
    var eRows = eventSheet.getDataRange().getValues();
    for (var er = 1; er < eRows.length; er++) {
      var en = eRows[er][0] ? eRows[er][0].toString().trim() : "";
      var ep = eRows[er][5] ? eRows[er][5].toString().trim() : "";
      if (en && ep && ep.indexOf("http") === 0) photoMap[en] = ep;
    }
  }

  var fallbackImg = "data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%221%22%20height%3D%221%22%3E%3Crect%20width%3D%221%22%20height%3D%221%22%20fill%3D%22%23111%22%2F%3E%3C%2Fsvg%3E";

  var rows = "";
  var totalYen = 0;
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var name = item.name + (item.sub ? " (" + item.sub + ")" : "");
    var qty  = item.qty || 0;
    var yen  = (item.yenEach || 0) * qty;
    totalYen += yen;
    var imgUrl = photoMap[item.name] || fallbackImg;
    rows += "<tr style=\"border-bottom:1px solid #eee;\">" +
      "<td style=\"padding:10px; width:60px;\">" +
        "<img src=\"" + imgUrl + "\" width=\"55\" height=\"55\" style=\"object-fit:cover; border-radius:4px; border:1px solid #eee;\" />" +
      "</td>" +
      "<td style=\"padding:10px; font-size:14px; color:#111;\">" + name + "</td>" +
      "<td style=\"padding:10px; font-size:16px; font-weight:700; color:#111; text-align:center;\">" + qty + "</td>" +
      "<td style=\"padding:10px; font-size:13px; color:#555; text-align:right;\">¥" + yen.toLocaleString() + "</td>" +
    "</tr>";
  }

  var html = "<div style=\"font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width:500px; margin:0 auto; color:#111;\">" +
    "<h2 style=\"font-size:17px; font-weight:600; border-bottom:1px solid #eee; padding-bottom:10px; margin-bottom:16px;\">" +
      "🛒 " + eventName + " — 待購清單" +
    "</h2>" +
    "<table style=\"width:100%; border-collapse:collapse;\">" +
      "<thead><tr style=\"background:#f5f5f5;\">" +
        "<th style=\"padding:8px 10px; font-size:11px; color:#888; text-align:left;\">&nbsp;</th>" +
        "<th style=\"padding:8px 10px; font-size:11px; color:#888; text-align:left;\">商品</th>" +
        "<th style=\"padding:8px 10px; font-size:11px; color:#888; text-align:center;\">數量</th>" +
        "<th style=\"padding:8px 10px; font-size:11px; color:#888; text-align:right;\">日幣</th>" +
      "</tr></thead>" +
      "<tbody>" + rows + "</tbody>" +
      "<tfoot><tr>" +
        "<td colspan=\"3\" style=\"padding:10px; font-size:13px; color:#888;\">合計</td>" +
        "<td style=\"padding:10px; font-size:15px; font-weight:700; text-align:right;\">¥" + totalYen.toLocaleString() + "</td>" +
      "</tr></tfoot>" +
    "</table>" +
    "<p style=\"font-size:11px; color:#aaa; margin-top:20px;\">852hk.81jp 後台自動發送</p>" +
  "</div>";

  MailApp.sendEmail({
    to: adminEmail,
    subject: "[待購清單] " + eventName + " — " + items.length + " 款商品",
    body: eventName + " 待購清單，共 " + items.length + " 款。",
    htmlBody: html
  });

  return { result: "success", to: adminEmail, count: items.length };
}

// =================================================================
// sendDeadlineReminderEmail：截單前 24 小時提醒（TW/HK 分流）
// =================================================================
function sendDeadlineReminderEmail(isTW, data) {
  if (!data.custEmail || data.custEmail.indexOf("@") === -1) return false;

  var parsedItems = parseSummaryItems(data.summary || "");
  var fallbackImg = "data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%221%22%20height%3D%221%22%3E%3Crect%20width%3D%221%22%20height%3D%221%22%20fill%3D%22%23111%22%2F%3E%3C%2Fsvg%3E";

  var greeting   = isTW ? "您好！" : "你好！";
  var payLabel   = isTW ? "匯款金額" : "付款金額";
  var currency   = isTW ? "NT$" : "HK$";
  var urgentNote = "⏰ 請於截單日期 <strong>" + data.deadlineStr + "</strong> 前完成付款，逾期未收到付款，訂單將自動取消。";
  var alreadyPaid = isTW
    ? "如果您已完成匯款，請直接回覆此 Email 或於 <a href='https://www.threads.com/@886tw.81jp?igshid=NTc4MTIwNjQ2YQ==' style='color:#2980b9;'>Threads</a> / <a href='https://www.instagram.com/886tw.81jp?igsh=MW8zMmVncGVwNmd1dg%3D%3D&utm_source=qr' style='color:#e1306c;'>Instagram</a> 私訊我們查詢，我們將盡快確認。"
    : "如果你已經完成付款，請直接回覆此 Email 或喺 <a href='https://www.threads.net/@852hk.81jp' style='color:#111;'>Threads</a> inbox 我哋查詢，我哋會盡快確認，謝謝！";

  // 進度條（待付款節點）
  var progressBar = buildProgressBar("pending");

  // 商品明細表格
  var tableRows = "";
  for (var k = 0; k < parsedItems.length; k++) {
    var it = parsedItems[k];
    var nameDisp = it.name + (it.sub ? " (" + it.sub + ")" : "");
    tableRows += "<tr style='border-bottom:1px solid #eee;'>" +
      "<td style='padding:10px 14px; font-size:13px; color:#333;'>" + nameDisp + "</td>" +
      "<td style='padding:10px 14px; font-size:13px; color:#555; text-align:center;'>" + it.qty + "</td>" +
      "</tr>";
  }

  // 付款資料區塊
  var payBlock = "";
  if (isTW) {
    var twPhone = data.custPhone ? data.custPhone.toString() : "";
    var last3   = twPhone.length >= 3 ? twPhone.slice(-3) : twPhone;
    payBlock = "<div style='background:#f0f7ff; border-left:4px solid #2980b9; padding:14px; border-radius:4px; margin:20px 0;'>" +
      "<p style='color:#1565c0; font-size:14px; font-weight:700; margin:0 0 10px;'>💰 郵局銀行匯款資料</p>" +
      "<p style='margin:6px 0 10px; font-size:15px; color:#1a237e;'>匯款金額：<strong style='font-size:20px;'>NT$ " + data.totalAmount + "</strong></p>" +
      "<p style='font-size:13px; color:#333; margin:4px 0;'>銀行：<strong>中華郵政（郵局）</strong></p>" +
      "<p style='font-size:13px; color:#333; margin:4px 0;'>局號／帳號：<strong>0041860-0025565</strong></p>" +
      "<p style='font-size:13px; color:#333; margin:4px 0;'>戶名：<strong>周◯恩</strong></p>" +
      "<p style='margin:14px 0 6px; font-size:14px; font-weight:700; color:#1565c0;'>匯款時請於備註欄填寫您的手機末3碼：</p>" +
      "<span style='display:inline-block; background:#fff; border:2px dashed #2980b9; padding:6px 18px; border-radius:6px; font-family:monospace; font-size:22px; font-weight:bold; color:#2980b9; margin:4px 0 10px; letter-spacing:6px;'>" + last3 + "</span>" +
      "<p style='margin:8px 0 0; font-size:12px; color:#888; line-height:1.7;'>匯款後如24小時內沒有收到已付款通知，請主動回覆電郵或<a href='https://www.threads.net/@852hk.81jp' style='color:#2980b9; font-weight:600;'>私訊我們</a>。</p>" +
      "</div>";
  } else {
    var payMethod = data.payMethod ? data.payMethod.toString() : "";
    if (payMethod.indexOf("PayMe") !== -1) {
      payBlock = "<div style='background:#fff0f0; border-left:4px solid #ff4d4d; padding:14px; border-radius:4px; margin:20px 0;'>" +
        "<p style='color:#d93838; font-size:14px; font-weight:700; margin:0 0 10px;'>💰 PayMe 付款資料</p>" +
        "<p style='margin:6px 0 10px; font-size:15px; color:#b71c1c;'>匯款金額：<strong style='font-size:20px;'>HK$ " + data.totalAmount + "</strong></p>" +
        "<p style='font-size:13px; color:#333; margin:8px 0;'><a href='https://payme.hsbc/miru' style='color:#ff4d4d; font-weight:bold;'>一按即 PayMe！ https://payme.hsbc/miru</a></p>" +
        "<p style='margin:4px 0 0; font-size:13px; font-weight:700; color:#c62828;'>付款時請喺【備註】填寫訂單編號：" + data.orderId + "</p>" +
        "</div>";
    } else {
      payBlock = "<div style='background:#f0f7ff; border-left:4px solid #0066cc; padding:14px; border-radius:4px; margin:20px 0;'>" +
        "<p style='color:#0052a3; font-size:14px; font-weight:700; margin:0 0 10px;'>💰 轉數快 (FPS) 付款資料</p>" +
        "<p style='margin:6px 0 10px; font-size:15px; color:#1a237e;'>匯款金額：<strong style='font-size:20px;'>HK$ " + data.totalAmount + "</strong></p>" +
        "<p style='font-size:13px; color:#333; margin:4px 0;'>轉數快號碼：<strong>8890873</strong></p>" +
        "<p style='font-size:13px; color:#333; margin:4px 0;'>帳戶持有人：<strong>Chow W. Y.</strong></p>" +
        "<p style='margin:4px 0 0; font-size:13px; font-weight:700; color:#0052a3;'>付款時請喺【備註】填寫訂單編號：" + data.orderId + "</p>" +
        "</div>";
    }
  }

  var subject = isTW
    ? "【截單提醒】" + data.eventName + " 即將截止，請盡快完成付款"
    : "【截單提醒】" + data.eventName + " 快截單喇，請盡快完成付款";

  var htmlBody =
    "<div style='font-family:Helvetica Neue,Helvetica,Arial,sans-serif; max-width:560px; margin:0 auto; color:#1a1a1a;'>" +
    "<div style='background:#111; color:#fff; text-align:center; padding:20px; border-radius:10px 10px 0 0;'>" +
      "<div style='font-size:20px; font-weight:700; letter-spacing:2px;'>852hk.81jp</div>" +
      "<div style='font-size:11px; color:#aaa; margin-top:4px;'>現場快閃代購</div>" +
    "</div>" +
    "<div style='background:#fff8e1; border:2px solid #ffc107; text-align:center; padding:12px; font-size:14px; font-weight:700; color:#e65100;'>" +
      "⏰ 截單前最後提醒 ⏰" +
    "</div>" +
    "<div style='padding:24px 20px;'>" +
      "<p>" + greeting + " " + (data.custName || "") + "</p>" +
      "<p>" + urgentNote + "</p>" +
      progressBar +
      "<p style='font-size:13px; color:#333;'>您的訂單 <strong>" + data.orderId + "</strong> 仍未完成付款：</p>" +
      "<table style='width:100%; border-collapse:collapse; margin:12px 0;'>" +
        "<thead><tr style='background:#f5f5f5;'>" +
          "<th style='padding:8px 14px; font-size:11px; color:#888; text-align:left;'>商品</th>" +
          "<th style='padding:8px 14px; font-size:11px; color:#888; text-align:center;'>數量</th>" +
        "</tr></thead><tbody>" + tableRows + "</tbody>" +
      "</table>" +
      payBlock +
      "<p style='font-size:12px; color:#888;'>" + alreadyPaid + "</p>" +
    "</div>" +
    "<div style='background:#f5f5f5; text-align:center; padding:12px; font-size:11px; color:#aaa; border-radius:0 0 10px 10px;'>" +
      "852hk.81jp | <a href='https://www.threads.net/@852hk.81jp' style='color:#888;'>Threads</a>" +
    "</div>" +
    "</div>";

  try {
    MailApp.sendEmail({
      to: data.custEmail,
      subject: subject,
      body: greeting + " 您的訂單 " + data.orderId + " 即將截單，請盡快完成付款。截止時間：" + data.deadlineStr,
      htmlBody: htmlBody
    });
    return true;
  } catch(e) {
    Logger.log("截單提醒發送失敗: " + e.toString());
    return false;
  }
}

// =================================================================
// sendStockoutInquiryEmail：缺貨詢問通知（TW/HK 分流）
// outOfStockItems, inStockItems: [{ name, sub, qty, price, photo }]
// =================================================================
function sendStockoutInquiryEmail(isTW, data) {
  if (!data.custEmail || data.custEmail.indexOf("@") === -1) return false;

  var greeting = isTW ? "您好！" : "你好！";
  var eventLabel = data.eventName || "代購活動";

  // ── 建立商品行 ──
  function makeItemRows(items, color) {
    var rows = "";
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      var name = it.name + (it.sub ? "（" + it.sub + "）" : "");
      var priceLabel = isTW ? "NT$ " : "HK$ ";
      var imgStyle = "width:50px;height:50px;object-fit:cover;border-radius:6px;border:1px solid #eee;display:block;";
      var imgTag = (it.photo && it.photo.indexOf("http") === 0)
        ? "<img src='" + it.photo + "' style='" + imgStyle + "' />"
        : "<div style='width:50px;height:50px;background:#eee;border-radius:6px;'></div>";
      rows +=
        "<tr><td style='padding:8px 12px; vertical-align:middle; width:60px;'>" + imgTag + "</td>" +
        "<td style='padding:8px 12px; font-size:14px; font-weight:600; color:#111; vertical-align:middle;'>" + name + "</td>" +
        "<td style='padding:8px 12px; font-size:12px; color:#888; vertical-align:middle; white-space:nowrap;'>" +
          priceLabel + (it.price || "") + " × " + (it.qty || 1) + "件" +
        "</td></tr>";
    }
    return rows;
  }

  var oos    = data.outOfStockItems  || [];
  var inStock = data.inStockItems    || [];

  // ── 缺貨框 ──
  var oosBlock = "";
  if (oos.length > 0) {
    oosBlock =
      "<div style='background:#fff3e0; border:1px solid #ffcc80; border-radius:8px; padding:14px 16px; margin:16px 0;'>" +
        "<div style='font-size:12px; font-weight:700; color:#e65100; margin-bottom:10px; letter-spacing:1px;'>⚠️ 以下商品尚未購入</div>" +
        "<table style='width:100%; border-collapse:collapse;'>" + makeItemRows(oos, "#fff3e0") + "</table>" +
      "</div>";
  }

  // ── 已購入框 ──
  var inStockBlock = "";
  if (inStock.length > 0) {
    inStockBlock =
      "<div style='background:#e8f5e9; border:1px solid #a5d6a7; border-radius:8px; padding:14px 16px; margin:16px 0;'>" +
        "<div style='font-size:12px; font-weight:700; color:#2e7d32; margin-bottom:10px; letter-spacing:1px;'>✅ 以下商品已成功購入</div>" +
        "<table style='width:100%; border-collapse:collapse;'>" + makeItemRows(inStock, "#e8f5e9") + "</table>" +
      "</div>";
  }

  // 計算缺貨商品退款金額（必須在 option1Detail 之前）
  var oosRefundTotal = 0;
  for (var ri = 0; ri < oos.length; ri++) {
    oosRefundTotal += (oos[ri].price || 0) * (oos[ri].qty || 1);
  }
  var oosRefundLabel = (isTW ? "NT$ " : "HK$ ") + oosRefundTotal.toLocaleString();

  // ── 退款選項（TW / HK 不同）──
  var option1Detail = isTW
    ?  "<div style='font-size:13px;color:#555;margin-bottom:6px;'>請複製以下資料，填寫後回覆此 Email：</div>" +
       "<pre style='background:#f8f8f8;border:1.5px solid #ddd;border-radius:8px;padding:14px;font-family:monospace;font-size:13px;line-height:2.4;white-space:pre;margin:4px 0;'>" +
        "💰 " + data.orderId + " 退款資料\n" +
        "退款金額： " + (isTW ? "NT$" : "HK$") + " " + oosRefundTotal + "\n" +
        "🏦 匯款銀行代碼：\n" +
        "🏦 銀行簡稱：\n" +
        "💳 銀行帳號：\n" +
        "👤 受款人戶名（如：陳〇文）：" +
       "</pre>"
    :  "<div style='background:#f5f5f5; border-radius:8px; padding:12px 14px; margin-top:10px; font-size:13px; color:#333; line-height:2.2;'>" +
        "💰 <strong>PayMe 退款</strong>：提供您嘅 PayMe 手提電話號碼<br>" +
        "🏦 <strong>FPS 退款</strong>：提供您嘅 FPS 手提電話號碼 / 快速支付號碼" +
       "</div>";



  var option1Intro = isTW
    ? "我們將對缺貨商品的金額（" + oosRefundLabel + "）進行退款。<br>請在回覆時提供以下退款資料："
    : "我們將對缺貨商品的金額（" + oosRefundLabel + "）進行退款。<br>請提供以下退款方式：";

  // 已購入商品區塊（若有）
  var siInSection = "";
  if (inStock.length > 0) {
    var siInRows = makeItemRows(siIn, "#e8f5e9");
    var siNote = isTW && data.shipmentRef
      ? "<div style='font-size:12px;color:#388e3c;margin-top:8px;'>📦 已成功購入商品將透過賣貨便（" + data.shipmentRef + "）寄出，運費約 NT$38，取貨付款。</div>"
      : "";
    siInSection =
      "<div style='background:#e8f5e9;border-left:4px solid #4caf50;border-radius:4px;padding:12px 14px;margin:14px 0;'>" +
        "<div style='font-size:13px;font-weight:700;color:#2e7d32;margin-bottom:8px;'>✅ 已成功購入商品</div>" +
        siInRows + siNote +
      "</div>";
  }

  var subject = (data.isResend ? "【再次通知】" : "【缺貨通知】") + (isTW ? "886tw.81jp" : "852hk.81jp") + " × " + (data.eventName || "") + " 部分商品缺貨";

  var htmlBody =
    "<div style='font-family:Helvetica Neue,Helvetica,Arial,sans-serif;max-width:560px;margin:0 auto;color:#1a1a1a;'>" +
    "<div style='background:#111;color:#fff;text-align:center;padding:20px 24px;'>" +
      "<div style='font-size:20px;font-weight:700;letter-spacing:2px;'>" + (isTW ? "886tw.81jp" : "852hk.81jp") + "</div>" +
      "<div style='font-size:11px;color:#aaa;margin-top:4px;'>現場快閃代購</div>" +
    "</div>" +
    "<div style='background:linear-gradient(135deg,#e65100,#f57c00);color:#fff;text-align:center;padding:14px 16px;'>" +
      "<div style='font-size:18px;font-weight:700;margin-bottom:4px;'>⚠️ 缺貨通知</div>" +
      "<div style='font-size:13px;opacity:0.9;'>部分商品未能購入，請查看詳情</div>" +
    "</div>" +
    "<div style='padding:24px;background:#fff;'>" +
      "<p style='font-size:15px;margin:0 0 16px;'>您好！" + (data.custName || "") + "</p>" +
      "<p style='font-size:14px;color:#333;line-height:1.8;margin:0 0 16px;'>" +
        "感謝您的訂購，但以下商品因現場已售罄，未能成功購入，非常抱歉。" +
      "</p>" +
      "<div style='background:#fff3e0;border-left:4px solid #ff9800;border-radius:4px;padding:12px 14px;margin:14px 0;'>" +
        "<div style='font-size:13px;font-weight:700;color:#e65100;margin-bottom:8px;'>⚠️ 以下商品未能購入</div>" +
        makeItemRows(oos, "#fff3e0") +
      "</div>" +
      siInSection +
      "<div style='background:#f9f9f9;border:1px solid #eee;border-radius:8px;padding:14px 16px;margin:16px 0;'>" +
        "<div style='font-size:13px;font-weight:700;color:#333;margin-bottom:8px;'>💰 退款資料</div>" +
        "<div style='font-size:13px;color:#555;line-height:1.7;margin-bottom:8px;'>" + option1Intro + "</div>" +
        option1Detail +
      "</div>" +
      "<p style='font-size:12px;color:#aaa;line-height:1.7;margin-top:16px;'>" +
        "如有任何問題，歡迎直接回覆此 Email 或私訊我們查詢，謝謝！</p>" + buildEmailFooter(isTW, "") + "<div style='display:none;'>Threads</a> / IG 私訊我們查詢。" +
      "</p>" +
    "</div>" +
    "<div style='background:#f5f5f5;text-align:center;padding:12px;font-size:11px;color:#aaa;'>886tw.81jpp</div>" +
    "</div>";

  try {
    MailApp.sendEmail({
      to: data.custEmail,
      subject: subject,
      body: "您好，" + (data.custName||"") + "，部分商品缺貨，退款金額 " + oosRefundLabel + "，詳情請查收 Email。",
      htmlBody: htmlBody
    });
    return true;
  } catch(e) {
    Logger.log("缺貨詢問發送失敗: " + e.toString());
    return false;
  }
}

// =================================================================
// sendArrivalEmail：到貨通知（TW）
// =================================================================
function sendArrivalEmail(ss, data) {
  if (!data.custEmail || data.custEmail.indexOf("@") === -1) return false;
  var parsedItems = parseSummaryItems(data.summary || "");
  var evSheet = getSheetByEventName(ss, data.eventName || "");
  var photoMap2 = {};
  if (evSheet) {
    var evR = evSheet.getDataRange().getValues();
    for (var ep = 1; ep < evR.length; ep++) {
      var en = evR[ep][0] ? evR[ep][0].toString().trim() : "";
      var ef = evR[ep][5] ? evR[ep][5].toString().trim() : "";
      if (en) photoMap2[en] = ef;
    }
  }
  var itemRowsHtml = "";
  for (var k = 0; k < parsedItems.length; k++) {
    var it = parsedItems[k];
    var nameDisp = it.name + (it.sub ? "（" + it.sub + "）" : "");
    var photo = findPhotoByName(photoMap2, it.name);
    var imgTag = (photo && photo.indexOf("http") === 0)
      ? "<img src='" + photo + "' style='width:56px;height:56px;object-fit:cover;border-radius:7px;border:1px solid #eee;display:block;' />"
      : "<div style='width:56px;height:56px;background:#f5f5f5;border-radius:7px;'></div>";
    itemRowsHtml +=
      "<tr style='border-bottom:1px solid #f0f0f0;'>" +
        "<td style='padding:10px 12px;width:68px;vertical-align:middle;'>" + imgTag + "</td>" +
        "<td style='padding:10px 12px;font-size:14px;font-weight:600;color:#111;vertical-align:middle;'>" + nameDisp + "</td>" +
        "<td style='padding:10px 12px;font-size:13px;color:#666;vertical-align:middle;white-space:nowrap;'>× " + it.qty + " 件</td>" +
      "</tr>";
  }
  var arrivalImgHtml = (data.arrivalPhotoUrl && data.arrivalPhotoUrl.indexOf("http") === 0)
    ? "<div style='flex-shrink:0;text-align:center;'><div style='font-size:10px;color:#059669;font-weight:700;margin-bottom:6px;'>到貨圖片</div><img src='" + data.arrivalPhotoUrl + "' style='width:90px;height:90px;object-fit:cover;border-radius:8px;border:2px solid #a7f3d0;display:block;' /></div>"
    : "";
  var subject = "【到貨通知】886tw.81jp × " + (data.eventName || "代購") + " 您的商品已到達台灣 🎉";
  var htmlBody =
    "<div style='font-family:Helvetica Neue,Helvetica,Arial,sans-serif;max-width:560px;margin:0 auto;color:#1a1a1a;'>" +
    "<div style='background:#111;color:#fff;text-align:center;padding:20px 24px;'><div style='font-size:20px;font-weight:700;letter-spacing:2px;'>852hk.81jp</div><div style='font-size:11px;color:#aaa;margin-top:4px;'>現場快閃代購</div></div>" +
    "<div style='background:linear-gradient(135deg,#10b981,#059669);color:#fff;text-align:center;padding:14px 16px;'><div style='font-size:18px;font-weight:700;margin-bottom:4px;'>🎉 您的商品已到達台灣！</div><div style='font-size:13px;opacity:0.9;'>請至賣貨便完成下單，取貨時再付款</div></div>" +
    "<div style='padding:24px;background:#fff;'>" +
      "<p style='font-size:15px;margin:0 0 16px;'>您好！" + (data.custName||"") + "</p>" +
      "<p style='font-size:14px;color:#333;line-height:1.8;margin:0 0 20px;'>感謝您訂購 <strong>886tw.81jp × " + (data.eventName||"") + "</strong> 的代購商品。<br>您的商品已安全抵達台灣，請點擊下方連結前往賣貨便賣場下單。<br>下單後將於 <strong>1-3 個工作天</strong>內安排寄出，謝謝！</p>" +
      "<div style='text-align:center;margin:20px 0;'><a href='" + (data.shopUrl||"#") + "' style='display:inline-block;background:#10b981;color:#fff;font-size:16px;font-weight:700;padding:14px 32px;border-radius:10px;text-decoration:none;'>📦 前往賣貨便下單</a><div style='font-size:11px;color:#aaa;margin-top:8px;'>" + (data.shopUrl||"") + "</div></div>" +
      "<hr style='border:none;border-top:1px dashed #e5e7eb;margin:20px 0;' />" +
      "<div style='margin:20px 0;padding:0 10px;'>" +"<div style='position:relative;display:flex;justify-content:space-between;align-items:center;width:100%;'>" +"<div style='position:absolute;top:5px;left:0;right:0;height:1px;background:#e5e5e5;z-index:1;'></div>" +"<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#2980b9;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #2980b9;'></div><div style='font-size:11px;font-weight:700;color:#2980b9;margin-top:7px;'>已收到訂單</div></div>" +"<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#2980b9;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #2980b9;'></div><div style='font-size:11px;font-weight:700;color:#2980b9;margin-top:7px;'>已付款</div></div>" +"<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#10b981;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #10b981;'></div><div style='font-size:11px;font-weight:700;color:#10b981;margin-top:7px;'>已到台</div></div>" +"<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#d1d5db;border-radius:50%;margin:0 auto;border:2px solid #fff;'></div><div style='font-size:11px;color:#888;margin-top:7px;'>已完成取貨</div></div>" +"</div></div>" +
      "<div style='font-size:12px;font-weight:700;color:#888;letter-spacing:1px;margin-bottom:12px;'>📋 訂單商品</div>" +
      "<table style='width:100%;border-collapse:collapse;'>" + itemRowsHtml + "</table>" +
      "<div style='background:#f8fffe;border:1px solid #a7f3d0;border-radius:10px;padding:14px 16px;margin:20px 0;display:flex;gap:14px;align-items:flex-start;'>" +
        arrivalImgHtml +
        "<div style='flex:1;'><div style='font-size:12px;font-weight:700;color:#059669;margin-bottom:10px;'>💰 賣貨便費用明細</div>" +
        "<div style='display:flex;justify-content:space-between;font-size:13px;color:#333;padding:4px 0;'><span style='color:#666;'>國際運費（" + (data.weight||0) + "g，最低50g計）</span><span style='font-weight:600;'>NT$ " + (data.shippingFee||0) + "</span></div>" +
        "<div style='display:flex;justify-content:space-between;font-size:13px;color:#333;padding:4px 0;'><span style='color:#666;'>賣貨便運費</span><span style='font-weight:600;'>NT$ 38</span></div>" +
        "<div style='height:1px;background:#d1fae5;margin:8px 0;'></div>" +
        "<div style='display:flex;justify-content:space-between;font-size:14px;color:#059669;font-weight:700;padding:4px 0;'><span>取貨時付款</span><span>NT$ " + ((data.shippingFee||0)+38) + "</span></div>" +
        "</div></div>" +
      "<div style='background:#fffbeb;border-left:4px solid #f59e0b;border-radius:4px;padding:12px 14px;margin:16px 0;'>" +
        "<div style='font-size:12px;font-weight:700;color:#92400e;margin-bottom:8px;'>⚠️ 取件注意事項</div>" +
        "<ul style='margin:0;padding-left:16px;font-size:13px;color:#555;line-height:2.2;'>" +
          "<li>賣貨便賣場費用為<strong>國際運費（NT$20／50g，不超過50g會收取NT$20，超過則以實重計算）＋ NT$38 賣貨便運費</strong>，取貨時付款。</li>" +
          "<li>收到 7-Eleven 的取貨簡訊後，請於<strong>簡訊指定的取貨期限內</strong>前往門市取貨。</li>" +
          "<li style='color:#dc2626;font-weight:600;'>若未在指定日期內取件，包裹將會退回，無法重新配送，需重新下單購買。</li>" +
        "</ul></div>" +
      "<div style='background:#f0f4ff;border:1px solid #c7d2fe;border-radius:8px;padding:12px 16px;margin:12px 0;font-size:12px;color:#4b5563;line-height:1.8;'>" +
        "📦 <strong>包裝說明：</strong>重量為商品淨重，已扣除底部乾淨墊底紙箱重量，會另外使用氣泡紙保護，及使用乾淨紙箱包裝寄出。" +
      "</div>" +
      "<p style='font-size:12px;color:#aaa;line-height:1.7;margin-top:16px;'>如有任何問題，歡迎直接回覆此 Email 或於 Threads / Instagram 私訊我們查詢，謝謝！</p>" + "<div style='text-align:center;margin-top:10px;'><a href='https://www.threads.com/@886tw.81jp?igshid=NTc4MTIwNjQ2YQ==' style='display:inline-block;background:#111;color:#fff;font-size:12px;font-weight:700;padding:6px 14px;border-radius:14px;text-decoration:none;margin:3px;'>Threads @886tw.81jp</a><a href='https://www.instagram.com/886tw.81jp?igsh=MW8zMmVncGVwNmd1dg%3D%3D&utm_source=qr' style='display:inline-block;background:linear-gradient(135deg,#f09433,#e6683c,#dc2743,#cc2366,#bc1888);color:#fff;font-size:12px;font-weight:700;padding:6px 14px;border-radius:14px;text-decoration:none;margin:3px;'>Instagram @886tw.81jp</a></div>" + "<div style='display:none;'><a href='https://www.threads.net/@852hk.81jp' style='color:#10b981;'>Threads</a> / IG 私訊我們查詢。</p>" +
    "</div>" +
    "<div style='background:#f5f5f5;text-align:center;padding:12px;font-size:11px;color:#aaa;'>852hk.81jp</div></div>";
  try {
    MailApp.sendEmail({ to: data.custEmail, subject: subject, body: "您好，商品已到台灣，請前往賣貨便下單：" + (data.shopUrl||""), htmlBody: htmlBody });
    return true;
  } catch(e) { Logger.log("到貨通知發送失敗: " + e.toString()); return false; }
}

// =================================================================
// sendRefundEmail：已退款通知（TW）
// =================================================================
function sendRefundEmail(ss, data) {
  if (!data.custEmail || data.custEmail.indexOf("@") === -1) return false;
  var eventName = data.eventName || "代購";
  var custName  = data.custName  || "";
  var refundAmt = data.stockoutAmt || 0;
  var refundPhoto = data.refundPhoto || "";
  var evSheet = getSheetByEventName(ss, eventName);
  var pricePhotoMap = {};
  if (evSheet) {
    var evRows = evSheet.getDataRange().getValues();
    for (var ep = 1; ep < evRows.length; ep++) {
      var epN = evRows[ep][0] ? evRows[ep][0].toString().trim() : "";
      var epS = evRows[ep][1] ? evRows[ep][1].toString().trim() : "";
      var epT = parseFloat(evRows[ep][8]) || 0;
      var epF = evRows[ep][5] ? evRows[ep][5].toString().trim() : "";
      if (epN) pricePhotoMap[epN + "|" + epS] = { price: epT, photo: epF };
    }
  }
  function lookupItem(name, sub) {
    var k1 = name + "|" + (sub || "");
    if (pricePhotoMap[k1]) return pricePhotoMap[k1];
    var k2 = name + "|";
    if (pricePhotoMap[k2]) return pricePhotoMap[k2];
    var bracketM = name.match(/^(.*?)\s*[(（]([^)）]+)[)）]\s*$/);
    if (bracketM) {
      var cn = bracketM[1].trim(), es = bracketM[2].trim();
      var k3 = cn + "|" + (sub||""); if (pricePhotoMap[k3]) return pricePhotoMap[k3];
      var k4 = cn + "|" + es;         if (pricePhotoMap[k4]) return pricePhotoMap[k4];
      var k5 = cn + "|";              if (pricePhotoMap[k5]) return pricePhotoMap[k5];
    }
    for (var fk in pricePhotoMap) { if (fk.split("|")[0] === name) return pricePhotoMap[fk]; }
    return { price: 0, photo: "" };
  }
  function makeRow(it, borderColor) {
    var info = lookupItem(it.name, it.sub);
    var nm = it.name + (it.sub ? "（" + it.sub + "）" : "");
    var imgHtml = (info.photo && info.photo.indexOf("http") === 0)
      ? "<img src='" + info.photo + "' style='width:52px;height:52px;object-fit:cover;border-radius:7px;border:1px solid #eee;flex-shrink:0;display:block;' />"
      : "<div style='width:52px;height:52px;background:#f0f0f0;border-radius:7px;flex-shrink:0;'></div>";
    var price = info.price * it.qty;
    return "<div style='display:flex;gap:10px;align-items:center;padding:8px 0;border-bottom:1px solid " + borderColor + ";'>" +
      imgHtml +
      "<div style='flex:1;'><div style='font-size:13px;font-weight:600;color:#111;'>" + nm + "</div><div style='font-size:12px;color:#888;margin-top:2px;'>× " + it.qty + " 件</div></div>" +
      "<div style='font-size:13px;font-weight:700;white-space:nowrap;'>NT$ " + (price > 0 ? price.toLocaleString() : "-") + "</div>" +
    "</div>";
  }
  var oosItems  = parseSummaryItems(data.stockoutItems || "");
  var allItems  = parseSummaryItems(data.summary || "");
  var oosKeys   = {};
  oosItems.forEach(function(it) { oosKeys[it.name + "|" + (it.sub||"")] = true; });
  var purchItems = allItems.filter(function(it) { return !oosKeys[it.name + "|" + (it.sub||"")]; });
  var oosRowsHtml  = oosItems.map(function(it) { return makeRow(it, "#fef2f2"); }).join("");
  var purchRowsHtml= purchItems.map(function(it) { return makeRow(it, "#dcfce7"); }).join("");
  var refundImgHtml = (refundPhoto && refundPhoto.indexOf("http") === 0)
    ? "<div style='text-align:center;margin:14px 0;'><img src='" + refundPhoto + "' style='max-width:240px;border-radius:10px;border:1px solid #bfdbfe;' /><div style='font-size:10px;color:#93c5fd;margin-top:4px;'>退款憑證</div></div>" : "";
  var purchasedSection = purchItems.length > 0
    ? "<div style='background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:12px 14px;margin:12px 0;'>" +
        "<div style='font-size:12px;font-weight:700;color:#16a34a;margin-bottom:8px;'>✅ 已成功購入商品</div>" +
        purchRowsHtml +
        "<div style='font-size:12px;color:#6b7280;margin-top:8px;padding-top:8px;border-top:1px solid #dcfce7;'>📦 商品到達台灣後，我們會另行發送賣貨便取件通知，請留意 Email。</div>" +
      "</div>" : "";
  var subject = "【退款通知】" + (typeof isTW !== "undefined" && isTW ? "886tw.81jp" : "852hk.81jp") + " × " + eventName + " 退款已完成";
  var htmlBody =
    "<div style='font-family:Helvetica Neue,Helvetica,Arial,sans-serif;max-width:560px;margin:0 auto;color:#1a1a1a;'>" +
    "<div style='background:#111;color:#fff;text-align:center;padding:20px 24px;'><div style='font-size:20px;font-weight:700;letter-spacing:2px;'>852hk.81jp</div><div style='font-size:11px;color:#aaa;margin-top:4px;'>現場快閃代購</div></div>" +
    "<div style='background:linear-gradient(135deg,#3b82f6,#1d4ed8);color:#fff;text-align:center;padding:14px 16px;'><div style='font-size:18px;font-weight:700;margin-bottom:4px;'>💸 退款通知</div><div style='font-size:13px;opacity:0.85;'>您的退款已完成</div></div>" +
    "<div style='padding:24px;background:#fff;'>" +
      "<p style='font-size:15px;margin:0 0 12px;'>您好！" + custName + "</p>" +
      "<p style='font-size:13px;color:#555;line-height:1.7;margin:0 0 16px;'>以下為您缺貨商品的退款明細，退款已完成，請查收。</p>" +
      refundImgHtml +
      "<div style='background:#fff5f5;border:1px solid #fecaca;border-radius:8px;padding:14px;margin:12px 0;'>" +
        "<div style='font-size:12px;font-weight:700;color:#dc2626;margin-bottom:10px;'>💸 退款商品</div>" +
        oosRowsHtml +
        "<div style='display:flex;justify-content:space-between;align-items:center;margin-top:10px;padding-top:10px;border-top:1px solid #fecaca;'>" +
          "<span style='font-size:13px;font-weight:600;color:#dc2626;'>退款總額</span>" +
          "<span style='font-size:18px;font-weight:800;color:#dc2626;'>NT$ " + refundAmt.toLocaleString() + "</span>" +
        "</div></div>" +
      purchasedSection +
      "<p style='font-size:12px;color:#aaa;line-height:1.7;margin-top:16px;'>如有任何問題，歡迎直接回覆此 Email 或於 <a href='https://www.threads.net/@852hk.81jp' style='color:#3b82f6;'>Threads</a> / IG 私訊我們查詢。</p>" +
    "</div>" +
    "<div style='background:#f5f5f5;text-align:center;padding:12px;font-size:11px;color:#aaa;'>852hk.81jp</div></div>";
  try {
    MailApp.sendEmail({ to: data.custEmail, subject: subject, body: "您好，" + custName + "，退款 NT$" + refundAmt + " 已完成。", htmlBody: htmlBody });
    return true;
  } catch(e) { Logger.log("退款通知發送失敗: " + e.toString()); return false; }
}

// =================================================================
// checkAndSetArrived：四欄齊全時將訂單狀態改為「已到貨」
// =================================================================
function checkAndSetArrived(sheet, rowIndex, row) {
  var status = row[15] ? row[15].toString().trim() : "";
  if (status !== "已付款" && status !== "已收款") return;
  var sVal = row[18] ? row[18].toString().trim() : "";
  var fVal = row[5]  ? parseFloat(row[5])  : 0;
  var wVal = row[22] ? parseFloat(row[22]) : 0;
  var xVal = row[23] ? row[23].toString().trim() : "";
  if (sVal && fVal > 0 && wVal > 0 && xVal) {
    sheet.getRange(rowIndex + 1, 16).setValue("已到貨");
    Logger.log("狀態改為已到貨：" + row[0]);
  }
}

function onSheetEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== "訂單紀錄") return;
  var col = e.range.getColumn();
  if ([6, 19, 23, 24].indexOf(col) === -1) return;
  var row = e.range.getRow();
  if (row < 2) return;
  var rowData = sheet.getRange(row, 1, 1, 24).getValues()[0];
  checkAndSetArrived(sheet, row - 1, rowData);
}


// =================================================================
// checkTakenEmails：掃描「買家取貨完成通知」→ 改狀態「已完結」
// Email 標題：賣貨便：買家取貨完成通知
// Body 含：CM號碼，從 Y欄比對
// =================================================================
function checkTakenEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();

  var props = PropertiesService.getScriptProperties();
  var doneStr = props.getProperty("takenEmailsDone") || "{}";
  var done;
  try { done = JSON.parse(doneStr); } catch(e) { done = {}; }
  var cutoff = Date.now() - 60 * 86400000;
  for (var pid in done) { if (done[pid] < cutoff) delete done[pid]; }

  // 搜尋多種可能格式：「買家取貨完成通知」或「買家完成取貨訂單通知」
  var threadSet = {};
  var allThreads = [];
  ["subject:買家取貨完成通知 newer_than:60d",
   "subject:買家完成取貨訂單通知 newer_than:60d",
   "subject:完成取貨 newer_than:60d"].forEach(function(q) {
    GmailApp.search(q, 0, 200).forEach(function(t) {
      if (!threadSet[t.getId()]) { threadSet[t.getId()] = true; allThreads.push(t); }
    });
  });
  var threads = allThreads;
  Logger.log("checkTakenEmails: 找到 " + threads.length + " 個 thread");

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var msgId = msg.getId();
      if (done[msgId]) continue;

      var html = msg.getBody() || "";
      var body = html
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<[^>]+>/g, " ")
        .replace(/&nbsp;/g, " ")
        .trim();
      var plain = msg.getPlainBody() || "";
      if (plain.trim().length > 50) body = plain;

      // 提取 CM 號碼
      var cmMatch = body.match(/(CM[\d]{6,})/);
      if (!cmMatch) { done[msgId] = Date.now(); continue; }
      var cmNumber = cmMatch[1];
      Logger.log("找到 CM: " + cmNumber);

      // 在 Y欄(index 24) 比對 CM 號
      var matched = false;
      for (var r = 1; r < rows.length; r++) {
        if (!rows[r][0]) continue;
        var rowCM = rows[r][24] ? rows[r][24].toString().trim() : "";
        if (rowCM === cmNumber) {
          sheet.getRange(r + 1, 16).setValue("已完結"); // P欄
          Logger.log("✅ 已完結：" + rows[r][0] + " CM:" + cmNumber);
          matched = true;
          // 自動發 TW 已取貨感謝 email
          try {
            if (rows[r][11] !== "HK") { // L欄 pickupType 非 HK
              sendTWCompletedEmail(ss, rows[r][0].toString());
            }
          } catch(teErr) { Logger.log("已取貨 email 錯誤: " + teErr); }
          break;
        }
      }
      if (!matched) Logger.log("⚠️ 找不到 CM: " + cmNumber);
      done[msgId] = Date.now();
    }
  }
  props.setProperty("takenEmailsDone", JSON.stringify(done));
}

// =================================================================
// checkRutenOrders：掃描賣貨便訂單成立郵件
// =================================================================

// =================================================================
// checkStoreCreatedEmails：掃描「賣貨便：賣場建立成功通知」
// 提取賣場網址 → 寫入訂單紀錄 S欄 (shipmentRef)
// =================================================================
function checkStoreCreatedEmails() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();

  var props = PropertiesService.getScriptProperties();
  var doneStr = props.getProperty("storeCreatedDone") || "{}";
  var done;
  try { done = JSON.parse(doneStr); } catch(e) { done = {}; }
  var cutoff = Date.now() - 30 * 86400000;
  for (var pid in done) { if (done[pid] < cutoff) delete done[pid]; }

  var threads = GmailApp.search("subject:賣場建立成功通知 newer_than:30d", 0, 200);
  Logger.log("checkStoreCreatedEmails: 找到 " + threads.length + " 個 thread");

  for (var t = 0; t < threads.length; t++) {
    var thread = threads[t];
    var messages = thread.getMessages();
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var msgId = msg.getId();
      if (done[msgId]) continue;

      // 解析 HTML body
      var html = msg.getBody() || msg.getPlainBody() || "";
      if (!html) { done[msgId] = Date.now(); continue; }

      // 轉純文字
      var body = html
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<\/td>/gi, " ")
        .replace(/<\/tr>/gi, "\n")
        .replace(/<[^>]+>/g, " ")
        .replace(/&nbsp;/g, " ")
        .replace(/&amp;/g, "&")
        .replace(/&lt;/g, "<")
        .replace(/&gt;/g, ">")
        .replace(/[ \t]{2,}/g, " ")
        .trim();

      // 只處理有「賣場_ORD」的 email
      if (body.indexOf("賣場_ORD") === -1) { done[msgId] = Date.now(); continue; }

      // 提取訂單號
      var ordMatch = body.match(/賣場_ORD([\d]{10,})/);
      if (!ordMatch) { done[msgId] = Date.now(); continue; }
      var orderId = "ORD" + ordMatch[1];

      // 提取賣場網址（myship.7-11.com.tw/general/detail/GM...）
      var urlMatch = html.match(/https?:\/\/myship\.7-11\.com\.tw\/general\/detail\/GM[\w]+/);
      if (!urlMatch) {
        // 也從純文字找
        urlMatch = body.match(/https?:\/\/myship\.7-11\.com\.tw\/general\/detail\/GM[\w]+/);
      }
      var shopUrl = urlMatch ? urlMatch[0] : "";

      Logger.log("訂單號: " + orderId + "  賣場網址: " + shopUrl);

      if (!shopUrl) { done[msgId] = Date.now(); continue; }

      // 比對訂單 A欄，寫入 S欄 (col 19)
      var matched = false;
      for (var r = 1; r < rows.length; r++) {
        if (!rows[r][0]) continue;
        if (rows[r][0].toString().trim() === orderId) {
          sheet.getRange(r + 1, 19).setValue(shopUrl); // S欄
          sheet.getRange(r + 1, 16).setValue("已到貨");   // P欄：狀態
          Logger.log("✅ 已更新 S欄 + 狀態已到貨：" + orderId + " → " + shopUrl);
          matched = true;
          break;
        }
      }
      if (!matched) Logger.log("⚠️ 找不到訂單: " + orderId);

      done[msgId] = Date.now();
    }
  }
  props.setProperty("storeCreatedDone", JSON.stringify(done));
}

// 手動清除已處理記錄，重新掃描時執行一次
function clearRutenDone() {
  PropertiesService.getScriptProperties().deleteProperty("rutenMsgDone");
  Logger.log("✅ rutenMsgDone 已清除，可重新掃描");
}

function checkRutenOrders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();

  // 用 message ID 追蹤（不用 thread ID），同一 thread 的新信也能處理
  var props = PropertiesService.getScriptProperties();
  var doneStr = props.getProperty("rutenMsgDone") || "{}";
  var done;
  try { done = JSON.parse(doneStr); } catch(e) { done = {}; }
  var cutoff = Date.now() - 30 * 86400000;
  for (var pid in done) { if (done[pid] < cutoff) delete done[pid]; }

  // 搜尋「賣貨便：訂單成立通知」（包括7-ELEVEN前綴）
  var threads = GmailApp.search("subject:訂單成立通知 newer_than:30d", 0, 200);
  Logger.log("checkRutenOrders: 找到 " + threads.length + " 個 thread");

  for (var t = 0; t < threads.length; t++) {
    var thread = threads[t];
    var messages = thread.getMessages();

    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var msgId = msg.getId();

      // 用 message ID 判斷是否已處理（而非 thread ID）
      if (done[msgId]) continue;

      // 解析 body
      var plainBody = msg.getPlainBody() || "";
      var body;
      if (plainBody.trim().length > 100) {
        body = plainBody;
      } else {
        var html = msg.getBody() || "";
        body = html
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<\/td>/gi, " ")
          .replace(/<\/tr>/gi, "\n")
          .replace(/<[^>]+>/g, " ")
          .replace(/&nbsp;/g, " ")
          .replace(/&amp;/g, "&")
          .replace(/&lt;/g, "<")
          .replace(/&gt;/g, ">")
          .replace(/[ \t]{2,}/g, " ")
          .trim();
      }
      if (!body) { done[msgId] = Date.now(); continue; }

      // 只處理有「賣場_ORD」的 email
      if (body.indexOf("賣場_ORD") === -1) {
        Logger.log("msg " + msgId + ": 非賣場_ORD，跳過");
        done[msgId] = Date.now();
        continue;
      }

      Logger.log("body片段: " + body.substring(0, 400).replace(/\n/g,"↵").replace(/\t/g,"→"));

      // 提取訂單號
      var orderIdRaw = "";
      var sm = body.match(/賣場_ORD([\d]{10,})/);
      if (sm) orderIdRaw = "ORD" + sm[1];
      if (!orderIdRaw) {
        Logger.log("msg " + msgId + ": 找不到訂單號");
        done[msgId] = Date.now();
        continue;
      }
      Logger.log("解析到訂單號: " + orderIdRaw);

      // CM 號
      var cmMatch = body.match(/訂單編號[：:\s]*(CM[\d]+)/);
      if (!cmMatch) cmMatch = body.match(/(CM[\d]{6,})/);
      var cmNumber = cmMatch ? cmMatch[1] : "";

      // 收件者姓名（收件者資訊：XXX 後面可能接空格或換行和門市）
      var recipientName = "";
      var rm = body.match(/收件者資訊[：:]\s*([^\s][^\n\r]{0,15})/);
      if (rm) { recipientName = rm[1].trim().split(/\s+/)[0].trim(); }
      if (!recipientName) {
        var rm2 = body.match(/([^\s]{1}[*＊][^\s]{1})/);
        if (rm2) recipientName = rm2[1].trim();
      }
      Logger.log("收件者: [" + recipientName + "]");

      // 門市名稱（在收件者資訊後面，連串格式也能抓到）
      var storeMatch = body.match(/收件者[\s\S]{0,200}?([\u4e00-\u9fff\w]+門市)/);
      var storeName = storeMatch ? storeMatch[1].trim() : "";
      Logger.log("門市: [" + storeName + "]  CM: [" + cmNumber + "]");

      // 比對訂單紀錄 A欄
      var matched = false;
      for (var r = 1; r < rows.length; r++) {
        if (!rows[r][0]) continue;
        var rowOrderId = rows[r][0].toString().trim();
        if (rowOrderId === orderIdRaw) {
          sheet.getRange(r + 1, 16).setValue("已下單");
          if (cmNumber)      sheet.getRange(r + 1, 25).setValue(cmNumber);
          if (recipientName) sheet.getRange(r + 1, 26).setValue(recipientName);
          if (storeName)     sheet.getRange(r + 1, 27).setValue(storeName);
          Logger.log("✅ 已更新訂單 " + rowOrderId);
          matched = true;
          break;
        }
      }
      if (!matched) Logger.log("⚠️ 找不到對應訂單: " + orderIdRaw);

      done[msgId] = Date.now(); // 標記這封 message 已處理
      thread.markRead();
    }
  }
  props.setProperty("rutenMsgDone", JSON.stringify(done));
}


// =================================================================
// checkRefundReplies：掃描客人回覆的退款銀行資料
// =================================================================
// =================================================================
// clearProcessed：手動清除已處理 thread 快取（需要時在 GAS 執行一次）
// =================================================================
function clearProcessed() {
  var p = PropertiesService.getScriptProperties();
  p.deleteProperty("refundRepliesProcessed");
  p.deleteProperty("refundMsgProcessed");
  Logger.log("✅ 已清除 refundRepliesProcessed 快取");
}


// =================================================================
// clearProcessed：清除已處理的 thread 記錄（需要重新掃描時手動執行）
// =================================================================


function checkRefundReplies() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();

  var props = PropertiesService.getScriptProperties();
  var processedKey = "refundMsgProcessed";  // 改用 message ID 追蹤
  var processedStr = props.getProperty(processedKey) || "{}";
  var processed;
  try { processed = JSON.parse(processedStr); } catch(e) { processed = {}; }
  var cutoff = Date.now() - 30 * 86400000;
  for (var pid in processed) { if (processed[pid] < cutoff) delete processed[pid]; }

  // 搜尋含退款資料的回覆（包括再次通知版本）
  var threadSet = {};
  var allThreads = [];
  ["subject:缺貨通知 newer_than:30d", "subject:再次通知 newer_than:30d"].forEach(function(q) {
    GmailApp.search(q, 0, 100).forEach(function(t) {
      if (!threadSet[t.getId()]) { threadSet[t.getId()] = true; allThreads.push(t); }
    });
  });
  var threads = allThreads;
  Logger.log("checkRefundReplies: 找到 " + threads.length + " 個 thread");

  for (var t = 0; t < threads.length; t++) {
    var thread = threads[t];
    var threadId = thread.getId();
    var messages = thread.getMessages();
    Logger.log("Thread " + threadId + " 共 " + messages.length + " 封");

    // 掃描所有郵件，找客人回覆（非 852hk 寄出的）
    for (var m = 0; m < messages.length; m++) {
      var msg = messages[m];
      var msgId = msg.getId();
      if (processed[msgId]) continue; // 已處理的 message 跳過
      var fromAddr = msg.getFrom() || "";
      // 跳過自己發的
      if (fromAddr.indexOf("852hk") !== -1 || fromAddr.indexOf("noreply") !== -1) { processed[msgId] = Date.now(); continue; }

      var plainBody = msg.getPlainBody() || "";
      var body = plainBody;
      // 只取客人新寫的內容（引用的原始 email 前），避免匹配到空模板
      var bodyParts = body.split(/\r?\n>\s/);
      var bodyNew = bodyParts[0] || body;
      // 如果分割後太短，取前20行
      if (bodyNew.length > body.length * 0.9) {
        bodyNew = body.split(/\r?\n/).slice(0, 20).join("\n");
      }
      body = bodyNew;
      // 若 plain text 沒有銀行資料（客人在 HTML 填寫），改用 HTML body 並移除 HTML tag
      if (body.indexOf("\u532f\u6b3e\u9280\u884c\u4ee3\u78bc") === -1 && body.indexOf("\u9000\u6b3e\u8cc7\u6599") === -1) {
        var htmlBody = msg.getBody() || "";
        // 移除 HTML tag，保留文字
        var strippedHtml = htmlBody.replace(/<br\s*\/?>/gi, "\n").replace(/<\/?(div|p|li|tr|td|th)[^>]*>/gi, "\n").replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
        // 同樣只取引用前的部分
        var htmlParts = strippedHtml.split(/\n>\s/);
        var htmlNew = htmlParts[0] || strippedHtml;
        if (htmlNew.indexOf("\u532f\u6b3e\u9280\u884c\u4ee3\u78bc") !== -1 || htmlNew.indexOf("\u9000\u6b3e\u8cc7\u6599") !== -1) {
          body = htmlNew;
          Logger.log("使用 HTML body 解析");
        }
      }
      Logger.log("From: " + fromAddr + " | body前100: " + body.substring(0, 100));

      // 確認含退款資料
      if (body.indexOf("退款資料") === -1 && body.indexOf("匯款銀行代碼") === -1) continue;

      // 解析訂單號：支援多種格式
      // 格式1: "TW20260601001 退款資料"
      // 格式2: "💰 TW20260601001 退款資料"
      // 解析訂單號：優先從主旨取，再從 body 取
      var msgSubject = msg.getSubject() || "";
      Logger.log("主旨: " + msgSubject);
      var orderId = "";
      // 從主旨找 ORD/TW/HK 格式
      var subjOrd = msgSubject.match(/ORD\d{10,}|(?:TW|HK)\d{8,}/);
      if (subjOrd) { orderId = subjOrd[0]; }
      // fallback: 從 body 找
      if (!orderId) {
        var bodyOrd = body.match(/ORD\d{10,}|(?:TW|HK)\d{8,}/);
        if (bodyOrd) { orderId = bodyOrd[0]; }
      }
      if (!orderId) { Logger.log("找不到訂單號，主旨: " + msgSubject); continue; }
      Logger.log("解析到訂單號: " + orderId);

      // 解析銀行資料（支援中文冒號和英文冒號）
      var bankNameMatch = body.match(/銀行簡稱\s*[：:]\s*([^\r\n]+)/);

      var bankCodeMatch = body.match(/匯款銀行代碼\s*[：:]\s*([^\r\n]+)/);

      var bankAccMatch  = body.match(/銀行帳號\s*[：:]\s*([^\r\n]+)/);

      // 受款人戶名：先匹配 ）：後的值（跳過示例），再 fallback 直接匹配 ：後
      // 受款人戶名：找該行最後一個冒號後的值（支援破損字元、有無示例括號）
      var accNameLine = body.match(/受款.{0,2}戶名[^\n]*/);
      var accName = "";
      if (accNameLine) {
        var ln = accNameLine[0];
        var lastColon = Math.max(ln.lastIndexOf("\u003a"), ln.lastIndexOf("\uff1a"));
        if (lastColon !== -1) accName = ln.substring(lastColon + 1).trim().replace(/_/g, "").trim();
      }

      var bankName = bankNameMatch ? bankNameMatch[1].trim().replace(/_/g, "").replace(/-/g, "").trim() : "";
      var bankCode = bankCodeMatch ? bankCodeMatch[1].trim().replace(/_/g, "").replace(/-/g, "").trim() : "";
      var bankAcc  = bankAccMatch  ? bankAccMatch[1].trim().replace(/_/g, "").replace(/-/g, "").trim() : "";

      Logger.log("銀行簡稱:" + bankName + " 代碼:" + bankCode + " 帳號:" + bankAcc + " 戶名:" + accName);

      if (!bankCode && !bankName && !bankAcc && !accName) continue;

      // 比對訂單並更新
      var found = false;
      for (var r = 1; r < rows.length; r++) {
        if (!rows[r][0]) continue;
        if (rows[r][0].toString().trim() === orderId) {
          if (bankName) sheet.getRange(r + 1, 30).setValue(bankName); // AD
          if (bankCode) sheet.getRange(r + 1, 31).setNumberFormat("@").setValue(bankCode); // AE
          if (bankAcc)  sheet.getRange(r + 1, 32).setNumberFormat("@").setValue(bankAcc);  // AF
          if (accName)  sheet.getRange(r + 1, 33).setValue(accName);  // AG
          Logger.log("✅ 退款銀行資料已儲存：" + orderId);
          found = true;
          break;
        }
      }
      if (found) {
        processed[msg.getId()] = Date.now(); // 用 message ID 追蹤
        break;
      }
    }
  }
  props.setProperty(processedKey, JSON.stringify(processed));
}

function checkAndSendReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();
  var now = new Date();

  for (var r = 1; r < rows.length; r++) {
    var orderId  = rows[r][0] ? rows[r][0].toString() : "";
    var date     = rows[r][1] ? new Date(rows[r][1]) : null;
    var evName   = rows[r][2] ? rows[r][2].toString().trim() : "";
    var summary  = rows[r][3] ? rows[r][3].toString() : "";
    var amt      = parseFloat(rows[r][4]) || 0;
    var fee      = parseFloat(rows[r][5]) || 0;
    var jpyAmt   = parseFloat(rows[r][6]) || 0;
    var payMethod= rows[r][7] ? rows[r][7].toString() : "";
    var custName = rows[r][8] ? rows[r][8].toString() : "";
    var phone    = rows[r][9] ? rows[r][9].toString() : "";
    var email    = rows[r][10] ? rows[r][10].toString().trim() : "";
    var pickupType= rows[r][11] ? rows[r][11].toString() : "";
    var pickupCode= rows[r][12] ? rows[r][12].toString() : "";
    var pickupName= rows[r][13] ? rows[r][13].toString() : "";
    var pickupAddr= rows[r][14] ? rows[r][14].toString() : "";
    var status   = rows[r][15] ? rows[r][15].toString().trim() : "";
    var remark   = rows[r][16] ? rows[r][16].toString() : "";
    var payTime  = rows[r][17] ? rows[r][17] : null;

    if (!orderId || !email || !date) continue;
    var isTW = pickupType.indexOf("賣貨便") !== -1 || payMethod.indexOf("郵局") !== -1;

    // 截止時間
    var deadlineStr = "";
    try {
      var evSheet = ss.getSheetByName(evName) || ss.getSheetByName(stripEndPrefix(evName));
      if (evSheet) {
        var dCell = evSheet.getRange("B1").getValue();
        if (dCell) deadlineStr = Utilities.formatDate(new Date(dCell), Session.getScriptTimeZone(), "MM/dd HH:mm");
      }
    } catch(e2) {}

    var data = {
      orderId: orderId, eventName: evName, custName: custName, phone: phone,
      custEmail: email, summary: summary, amount: amt, shippingFee: fee,
      jpyAmount: jpyAmt, payMethod: payMethod, pickupType: pickupType,
      pickupCode: pickupCode, pickupName: pickupName, pickupAddress: pickupAddr,
      deadline: deadlineStr
    };

    // ── 待處理訂單：超過20分鐘未付款 → 自動改狀態「已取消」（取消email不自動發）
    if (status === "待處理") {
      var minsSince = (now - date) / 60000;
      if (minsSince >= 20 && remark.indexOf("已自動取消") === -1) {
        sheet.getRange(r + 1, 16).setValue("已取消");
        sheet.getRange(r + 1, 17).setValue((remark ? remark + " | " : "") + "已自動取消(20分鐘未付款):" + Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd HH:mm"));
        Logger.log("⏱️ 自動取消訂單：" + orderId);
      }
    }

    // ── 截單前24小時提醒
    if (status === "待處理" && deadlineStr && remark.indexOf("已截單催款") === -1) {
      try {
        var evSheet2 = ss.getSheetByName(evName);
        if (evSheet2) {
          var dl = new Date(evSheet2.getRange("B1").getValue());
          var hoursToDeadline = (dl - now) / 3600000;
          if (hoursToDeadline > 0 && hoursToDeadline <= 24) {
            sendDeadlineReminderEmail(isTW, data);
            sheet.getRange(r + 1, 17).setValue(remark + " | 已截單催款:" + Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd HH:mm"));
          }
        }
      } catch(e3) {}
    }
  }
}

// =================================================================
// doPost：處理所有 POST 請求
// =================================================================

// =================================================================
// ── Supabase 整合 ──────────────────────────────────────────────
// =================================================================
// ── 品牌資訊 ──
var TW_BRAND = {
  name: "886tw.81jp",
  threads: "https://www.threads.com/@886tw.81jp?igshid=NTc4MTIwNjQ2YQ==",
  instagram: "https://www.instagram.com/886tw.81jp?igsh=MW8zMmVncGVwNmd1dg%3D%3D&utm_source=qr",
  handle: "@886tw.81jp",
  subject_suffix: "— 886tw.81jp 台灣"
};
var HK_BRAND = {
  name: "852hk.81jp",
  threads: "https://www.threads.net/@852hk.81jp",
  instagram: "",
  handle: "@852hk.81jp",
  subject_suffix: "- 852hk.81jp"
};

// 標準 email footer 生成
function buildEmailFooter(isTW, mainMsg) {
  var brand = isTW ? TW_BRAND : HK_BRAND;
  var socialHtml = isTW
    ? "<a href='" + brand.threads + "' style='display:inline-block;background:#111;color:#fff;font-size:12px;font-weight:700;padding:7px 16px;border-radius:16px;text-decoration:none;margin:3px;'>Threads " + brand.handle + "</a>" +
      "<a href='" + brand.instagram + "' style='display:inline-block;background:linear-gradient(135deg,#f09433,#e6683c,#dc2743,#cc2366,#bc1888);color:#fff;font-size:12px;font-weight:700;padding:7px 16px;border-radius:16px;text-decoration:none;margin:3px;'>Instagram " + brand.handle + "</a>"
    : "<a href='" + brand.threads + "' style='display:inline-block;background:#111;color:#fff;font-size:12px;font-weight:700;padding:7px 16px;border-radius:16px;text-decoration:none;margin:3px;'>Threads " + brand.handle + "</a>";

  return "<div style='margin-top:28px;padding-top:16px;border-top:1px dashed #ddd;font-size:12px;color:#777;line-height:1.9;'>" +
    (mainMsg || "") +
    "<div style='margin-top:14px;padding:14px 16px;background:#f5f5f5;border-radius:8px;text-align:center;'>" +
      "<div style='font-size:11px;color:#999;margin-bottom:10px;'>" + (isTW ? "追蹤我們，獲得最新代購資訊 👀" : "追蹤我哋，獲得最新代購資訊 👀") + "</div>" +
      socialHtml +
    "</div>" +
    "<p style='margin-top:12px;font-size:12px;color:#999;text-align:center;'><strong style='color:#333;font-size:13px;'>" + brand.name + "</strong></p>" +
    "</div>";
}


var SUPABASE_URL = "https://pksqfpirggvsftvqrtji.supabase.co";
var SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBrc3FmcGlyZ2d2c2Z0dnFydGppIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI0MDE3MDQsImV4cCI6MjA5Nzk3NzcwNH0.TyGBeMNn91UCNEVhT-q4ivtgvXQL_GOLmdvtGMUIWuc";

// 通用 Supabase REST 呼叫
function sbFetch(path, method, body) {
  var options = {
    method: method || "GET",
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": "Bearer " + SUPABASE_KEY,
      "Content-Type": "application/json",
      "Prefer": method === "POST" ? "resolution=merge-duplicates" : ""
    },
    muteHttpExceptions: true
  };
  if (body) options.payload = JSON.stringify(body);
  var res = UrlFetchApp.fetch(SUPABASE_URL + "/rest/v1/" + path, options);
  var text = res.getContentText();
  try { return text ? JSON.parse(text) : []; } catch(e) { return []; }
}

function sbUpsert(table, rows) {
  if (!rows || rows.length === 0) return;

  var options = {
    method: "POST",
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": "Bearer " + SUPABASE_KEY,
      "Content-Type": "application/json",
      "Prefer": "resolution=merge-duplicates,return=minimal"
    },
    payload: JSON.stringify(rows),
    muteHttpExceptions: true
  };

  var res = UrlFetchApp.fetch(SUPABASE_URL + "/rest/v1/" + table, options);
  var code = res.getResponseCode();
  var text = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error("Supabase upsert " + table + " failed: HTTP " + code + " - " + text);
  }
}

function sbDelete(table, filter) {
  var options = {
    method: "DELETE",
    headers: {
      "apikey": SUPABASE_KEY,
      "Authorization": "Bearer " + SUPABASE_KEY,
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };
  UrlFetchApp.fetch(SUPABASE_URL + "/rest/v1/" + table + "?" + filter, options);
}

// =================================================================
// syncEventToSupabase：同步單一活動的商品到 Supabase
// 在 admin 手動觸發，或上架商品後觸發
// =================================================================
function syncEventToSupabase(ss, eventName) {
  if (!eventName) return { result: "error", message: "需要活動名稱" };
  var cleanName = stripEndPrefix(eventName);
  var isEnded   = eventName.indexOf("[END]") === 0;

  // 讀取商品 sheet
  var sheet = getSheetByEventName(ss, eventName);
  if (!sheet) return { result: "error", message: "找不到活動 sheet: " + eventName };

  var rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return { result: "ok", message: "活動無商品" };

  // 讀取截單時間
  var dlSheet = ss.getSheetByName("收單截止時間");
  var deadline = null;
  if (dlSheet) {
    var dlRows = dlSheet.getDataRange().getValues();
    for (var di = 1; di < dlRows.length; di++) {
      if (dlRows[di][0] && dlRows[di][0].toString().trim() === cleanName) {
        var dlVal = dlRows[di][1];
        if (dlVal instanceof Date) deadline = dlVal.toISOString();
        break;
      }
    }
  }

  // 讀取區域設定（從 sheet 第 1 行 E 欄判斷，或 sheet 名稱判斷）
  // 預設 BOTH，如果有 [HK] 或 [TW] 標記就用那個
  var region = "BOTH";
  if (cleanName.toLowerCase().indexOf("hk") !== -1) region = "HK";
  if (cleanName.toLowerCase().indexOf("tw") !== -1) region = "TW";

  // Upsert event（用 name 做衝突鍵）
  sbUpsert("events?on_conflict=name", [{
    name: cleanName,
    deadline: deadline,
    is_active: !isEnded && (!deadline || new Date() < new Date(deadline)),
    is_ended: isEnded,
    updated_at: new Date().toISOString()
  }]);

  // 刪除舊商品（全部重建）
  sbDelete("product_subs", "event_name=eq." + encodeURIComponent(cleanName));
  sbDelete("products", "event_name=eq." + encodeURIComponent(cleanName));

  // 讀取商品資料
  // Sheet 欄位格式（來自 addNewProduct）：
  // A(0)=名稱, B(1)=款式(concatenated), C(2)=日幣, D(3)=成本HKD,
  // E(4)=售價HKD, F(5)=圖片, G(6)=重量g, H(7)=備註, I(8)=台幣, J(9)=限購
  var products = [];
  var subRows  = [];
  var sortIdx  = 0;

  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    var productName = row[0] ? row[0].toString().trim() : "";
    if (!productName || productName === "SubTotal") continue;

    var subsRaw   = row[1] ? row[1].toString() : "";       // B: 款式
    var yenPrice  = Number(row[2]) || 0;                   // C: 日幣
    // row[3] = 成本HKD (略過)
    var hkdPrice  = Number(row[4]) || 0;                   // E: 售價HKD
    var imageUrl  = row[5] ? row[5].toString().trim() : ""; // F: 圖片
    // row[6] = 重量g (略過)
    // row[7] = 備註 (略過)
    var twdPrice  = Number(row[8]) || 0;                   // I: 台幣
    var stockLimitRaw = row[9];                            // J: 限購

    products.push({
      event_name:  cleanName,
      name:        productName,
      image_url:   imageUrl,
      yen_price:   yenPrice,
      hkd_price:   hkdPrice,
      twd_price:   twdPrice,
      sort_order:  sortIdx++
    });

    // 解析款式（B欄以逗號或頓號分隔，例如：A款,B款 或 A款，B款）
    var stockLimits = parseSubStockLimit(stockLimitRaw);
    var subsArray = subsRaw.split(/[,，]/).map(function(s){ return s.trim(); }).filter(function(s){ return s; });

    if (subsArray.length === 0) {
      var limit = stockLimits ? (stockLimits["__all__"] !== undefined ? stockLimits["__all__"] : null) : null;
      subRows.push({
        event_name:   cleanName,
        product_name: productName,
        sub_name:     "",
        stock_limit:  limit,
        image_url:    imageUrl,
        sort_order:   0
      });
    } else {
      subsArray.forEach(function(subName, si) {
        var limit = null;
        if (stockLimits) {
          if (stockLimits.hasOwnProperty(subName)) limit = stockLimits[subName];
          else if (stockLimits["__all__"] !== undefined) limit = stockLimits["__all__"];
        }
        subRows.push({
          event_name:   cleanName,
          product_name: productName,
          sub_name:     subName,
          stock_limit:  limit,
          image_url:    imageUrl,
          sort_order:   si
        });
      });
    }
  }

  if (products.length > 0) sbUpsert("products", products);
  if (subRows.length > 0) sbUpsert("product_subs", subRows);

  Logger.log("[syncEvent] " + cleanName + ": " + products.length + " 商品, " + subRows.length + " 款式");
  return { result: "ok", synced: products.length };
}

// =================================================================
// syncAllActiveEventsToSupabase：同步所有現有活動
// =================================================================
function syncAllActiveEventsToSupabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var EXCL = [
    "訂單紀錄",
    "收單截止時間",
    "Blank",
    "購貨紀錄",
    "盈利紀錄",
    "易寄取地址",
    "[Data]購貨紀錄",
    "[Data]盈利紀錄"
  ];

  var sheets = ss.getSheets()
    .map(function(sh) { return sh.getName(); })
    .filter(function(name) {
      if (!name) return false;
      if (name.indexOf("[Data]") === 0) return false;
      if (name.indexOf("[END]") === 0) return false;
      if (EXCL.indexOf(name) !== -1) return false;
      return true;
    });

  var ok = 0;
  var failed = 0;
  var details = [];

  sheets.forEach(function(name) {
    try {
      var res = syncEventToSupabase(ss, name);
      if (res && res.result === "ok") {
        ok++;
      } else {
        failed++;
      }
      details.push({
        eventName: name,
        result: res && res.result ? res.result : "unknown",
        synced: res && res.synced ? res.synced : 0,
        message: res && res.message ? res.message : ""
      });
    } catch (err) {
      failed++;
      details.push({
        eventName: name,
        result: "error",
        synced: 0,
        message: err.toString()
      });
    }
  });

  return {
    result: failed ? "partial" : "ok",
    total: sheets.length,
    ok: ok,
    failed: failed,
    details: details
  };
}

// =================================================================
// syncOrderedQtyToSupabase：同步訂單數量到 Supabase（手動觸發）
// =================================================================
function syncOrderedQtyToSupabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return;

  var rows = sheet.getDataRange().getValues();
  var qtyMap = {}; // key: eventName|productName|subName

  for (var r = 1; r < rows.length; r++) {
    var status = rows[r][15] ? rows[r][15].toString().trim() : "";
    if (status === "已取消" || status === "已完結") continue;
    var evName  = rows[r][2]  ? rows[r][2].toString().trim()  : "";
    var summary = rows[r][3]  ? rows[r][3].toString()         : "";
    var isPaid  = !!rows[r][17]; // R欄 paymentTime
    if (!evName || !summary) continue;
    var cleanEv = stripEndPrefix(evName);

    var items = parseSummaryItems(summary);
    items.forEach(function(item) {
      var key = cleanEv + "|" + item.name + "|" + (item.sub || "");
      if (!qtyMap[key]) qtyMap[key] = { event_name: cleanEv, product_name: item.name, sub_name: item.sub || "", qty_ordered: 0, qty_paid: 0 };
      qtyMap[key].qty_ordered += item.qty;
      if (isPaid) qtyMap[key].qty_paid += item.qty;
    });
  }

  var upsertRows = Object.values(qtyMap);
  if (upsertRows.length > 0) {
    upsertRows.forEach(function(r) { r.updated_at = new Date().toISOString(); });
    sbUpsert("ordered_qty", upsertRows);
  }
  Logger.log("[syncOrderedQty] 更新 " + upsertRows.length + " 筆");
}

// =================================================================
// updateOrderedQtyInSupabase：下單時即時更新庫存（在 doPost 呼叫）
// =================================================================
function updateOrderedQtyInSupabase(eventName, summaryStr, isPaid) {
  try {
    var cleanEv = stripEndPrefix(eventName);
    var items = parseSummaryItems(summaryStr);
    items.forEach(function(item) {
      var filter = "event_name=eq." + encodeURIComponent(cleanEv) +
                   "&product_name=eq." + encodeURIComponent(item.name) +
                   "&sub_name=eq." + encodeURIComponent(item.sub || "");
      // 先讀現有值
      var existing = sbFetch("ordered_qty?" + filter, "GET");
      var current  = existing && existing[0] ? existing[0] : { qty_ordered: 0, qty_paid: 0 };
      sbUpsert("ordered_qty", [{
        event_name:   cleanEv,
        product_name: item.name,
        sub_name:     item.sub || "",
        qty_ordered:  (current.qty_ordered || 0) + item.qty,
        qty_paid:     (current.qty_paid    || 0) + (isPaid ? item.qty : 0),
        updated_at:   new Date().toISOString()
      }]);
    });
  } catch(e) {
    Logger.log("[updateOrderedQty] error: " + e.toString());
  }
}

// =================================================================
// syncHKAddressesToSupabase：同步易寄取地址
// =================================================================
function syncHKAddressesToSupabase() {
  // Sheet "易寄取地址" 欄位順序：
  // A(0)=類別 | B(1)=編號 | C(2)=名稱 | D(3)=地區 | E(4)=分區 | F(5)=地址
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("易寄取地址");
  if (!sheet) { Logger.log("[syncHKAddresses] 找不到「易寄取地址」sheet"); return; }
  var rows = sheet.getDataRange().getValues();
  var addrs = [];
  for (var r = 1; r < rows.length; r++) {
    if (!rows[r][1]) continue; // 編號為空則略過
    addrs.push({
      type:       rows[r][0] ? rows[r][0].toString().trim() : "",
      code:       rows[r][1].toString().trim(),
      name:       rows[r][2] ? rows[r][2].toString().trim() : "",
      region:     rows[r][3] ? rows[r][3].toString().trim() : "",
      district:   rows[r][4] ? rows[r][4].toString().trim() : "",
      address:    rows[r][5] ? rows[r][5].toString().trim() : "",
      sort_order: r - 1
    });
  }
  if (addrs.length > 0) sbUpsert("hk_addresses", addrs);
  Logger.log("[syncHKAddresses] 同步 " + addrs.length + " 個地址");
}

// =================================================================
// syncPurchaseRecordsToSupabase：同步買貨區到 Supabase
// =================================================================
function syncPurchaseRecordsToSupabase(eventName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cleanEv = stripEndPrefix(eventName || "");
  if (!cleanEv) return;

  var pSheet = getEventPurchaseSheet(ss, eventName, false);
  if (!pSheet) return;
  var rows = pSheet.getDataRange().getValues();
  if (rows.length < 2) return;

  // 刪除舊記錄
  sbDelete("purchase_records", "event_name=eq." + encodeURIComponent(cleanEv));
  sbDelete("purchase_batches", "event_name=eq." + encodeURIComponent(cleanEv));

  var records = [];
  var batches  = [];

  // 讀 header row (row 0) 找批次日期（H欄起）
  var headerRow = rows[0];
  var batchDates = [];
  for (var bc = 7; bc < headerRow.length; bc++) {
    var bDateRaw = headerRow[bc];
    var bDate = "";
    if (bDateRaw instanceof Date) {
      bDate = Utilities.formatDate(bDateRaw, Session.getScriptTimeZone(), "yyyy/MM/dd");
    } else if (bDateRaw) {
      bDate = bDateRaw.toString().trim();
    }
    if (!bDate || bDate === "SubTotal" || bDate.length < 6) { batchDates.push(null); continue; }
    batchDates.push(bDate);
  }

  // 讀商品行
  for (var r = 1; r < rows.length; r++) {
    var row = rows[r];
    if (!row[0] || row[0].toString().trim() === "SubTotal") continue;
    var pName = row[0].toString().trim();
    var sub   = row[1] ? row[1].toString().trim() : "";

    records.push({
      event_name:    cleanEv,
      product_name:  pName,
      sub_name:      sub,
      yen_price:     Number(row[2]) || 0,
      qty_paid:      Number(row[3]) || 0,
      qty_unpaid:    Number(row[4]) || 0,
      qty_purchased: Number(row[5]) || 0,
      qty_settled:   Number(row[6]) || 0,
      updated_at:    new Date().toISOString()
    });

    // 批次數量
    for (var bc = 0; bc < batchDates.length; bc++) {
      if (!batchDates[bc]) continue;
      var qty = Number(row[7 + bc]) || 0;
      if (qty <= 0) continue;
      batches.push({
        event_name:   cleanEv,
        batch_date:   batchDates[bc],
        product_name: pName,
        sub_name:     sub,
        qty:          qty
      });
    }
  }

  if (records.length > 0) sbUpsert("purchase_records", records);
  if (batches.length > 0) sbUpsert("purchase_batches", batches);
  Logger.log("[syncPurchaseRecords] " + cleanEv + ": " + records.length + " 商品, " + batches.length + " 批次");
}

// =================================================================
// doPost 攔截：下單時更新 ordered_qty
// 在現有 doPost 的 submitOrder 成功後呼叫
// =================================================================


// =================================================================
// updateQtyPaidInSupabase：付款確認時更新 qty_paid
// =================================================================
function updateQtyPaidInSupabase(cleanEventName, summaryStr, delta) {
  try {
    var items = parseSummaryItems(summaryStr);
    items.forEach(function(item) {
      var filter = "event_name=eq." + encodeURIComponent(cleanEventName) +
                   "&product_name=eq." + encodeURIComponent(item.name) +
                   "&sub_name=eq." + encodeURIComponent(item.sub || "");
      var existing = sbFetch("ordered_qty?" + filter, "GET");
      var cur = existing && existing[0] ? existing[0] : { qty_ordered: 0, qty_paid: 0 };
      sbUpsert("ordered_qty", [{
        event_name: cleanEventName, product_name: item.name, sub_name: item.sub || "",
        qty_ordered: cur.qty_ordered || 0,
        qty_paid: Math.max(0, (cur.qty_paid || 0) + delta * item.qty),
        updated_at: new Date().toISOString()
      }]);
    });
  } catch(e) { Logger.log("[updateQtyPaid] " + e.toString()); }
}

// =================================================================
// updateQtyOrderedInSupabase：取消訂單時退回庫存
// =================================================================
function updateQtyOrderedInSupabase(cleanEventName, summaryStr, orderedDelta, paidDelta) {
  try {
    var items = parseSummaryItems(summaryStr);
    items.forEach(function(item) {
      var filter = "event_name=eq." + encodeURIComponent(cleanEventName) +
                   "&product_name=eq." + encodeURIComponent(item.name) +
                   "&sub_name=eq." + encodeURIComponent(item.sub || "");
      var existing = sbFetch("ordered_qty?" + filter, "GET");
      var cur = existing && existing[0] ? existing[0] : { qty_ordered: 0, qty_paid: 0 };
      sbUpsert("ordered_qty", [{
        event_name: cleanEventName, product_name: item.name, sub_name: item.sub || "",
        qty_ordered: Math.max(0, (cur.qty_ordered || 0) + orderedDelta * item.qty),
        qty_paid:    Math.max(0, (cur.qty_paid    || 0) + paidDelta   * item.qty),
        updated_at:  new Date().toISOString()
      }]);
    });
  } catch(e) { Logger.log("[updateQtyOrdered] " + e.toString()); }
}



// =================================================================
// debugSupabaseSync：診斷 Supabase 同步狀況
// 在 GAS 執行，看執行記錄就能知道問題所在
// =================================================================
function debugSupabaseSync() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log("=== Step 1: 查 Supabase purchase_records ===");
  var sbRows = sbFetch("purchase_records?select=event_name,product_name,qty_paid,qty_purchased&limit=20", "GET");
  if (!sbRows || sbRows.length === 0) {
    Logger.log("❌ Supabase purchase_records 完全空白！");
  } else {
    Logger.log("✅ Supabase 有 " + sbRows.length + " 筆記錄");
    sbRows.slice(0,5).forEach(function(r) {
      Logger.log("  · " + r.event_name + " | " + r.product_name + " | paid=" + r.qty_paid + " purchased=" + r.qty_purchased);
    });
  }

  Logger.log("\n=== Step 2: 本地購貨紀錄 Sheets ===");
  var sheets = ss.getSheets();
  var purchaseSheets = sheets.filter(function(s){ return s.getName().indexOf("購貨紀錄") !== -1; });
  Logger.log("找到 " + purchaseSheets.length + " 個購貨紀錄 sheet");
  purchaseSheets.forEach(function(ps) {
    var rows = ps.getDataRange().getValues();
    Logger.log("  · " + ps.getName() + "：" + (rows.length-1) + " 行商品，" + ps.getLastColumn() + " 欄");
    if (rows.length > 1) {
      Logger.log("    row1(header)=" + JSON.stringify(rows[0].slice(0,8)));
      Logger.log("    row2(sample)=" + JSON.stringify(rows[1].slice(0,8)));
    }
  });

  Logger.log("\n=== Step 3: 嘗試直接 upsert 一筆測試資料 ===");
  var testRow = [{
    event_name: "TEST_DEBUG",
    product_name: "Test Product",
    sub_name: "",
    yen_price: 1000,
    qty_paid: 1,
    qty_unpaid: 0,
    qty_purchased: 0,
    qty_settled: 0,
    updated_at: new Date().toISOString()
  }];
  sbUpsert("purchase_records", testRow);
  var checkRows = sbFetch("purchase_records?event_name=eq.TEST_DEBUG&select=*", "GET");
  if (checkRows && checkRows.length > 0) {
    Logger.log("✅ 直接 upsert 成功！Supabase 寫入正常");
    // 清除測試資料
    sbDelete("purchase_records", "event_name=eq.TEST_DEBUG");
    Logger.log("   (測試資料已清除)");
  } else {
    Logger.log("❌ upsert 失敗！可能是 RLS 或 API key 問題");
  }

  Logger.log("\n=== Step 4: 重新嘗試同步第一個活動 ===");
  var EXCL = ["訂單紀錄","易寄取地址","Blank","收單截止時間","購貨紀錄","盈利紀錄","[Data]盈利紀錄"];
  var allNames = ss.getSheets().map(function(s){ return s.getName(); })
    .filter(function(n){ return !n.startsWith("[Data]") && EXCL.indexOf(n) === -1; });
  Logger.log("活動列表: " + JSON.stringify(allNames));

  if (allNames.length > 0) {
    var firstEv = allNames[0];
    Logger.log("嘗試同步: " + firstEv);
    var pSheet = getEventPurchaseSheet(ss, firstEv, false);
    if (!pSheet) {
      Logger.log("❌ 找不到 [Data](" + stripEndPrefix(firstEv) + ")購貨紀錄 sheet");
    } else {
      var pRows = pSheet.getDataRange().getValues();
      Logger.log("✅ 找到購貨紀錄 sheet，共 " + pRows.length + " 行");
      syncPurchaseRecordsToSupabase(firstEv);
      // 驗證
      var afterSync = sbFetch("purchase_records?event_name=eq." + encodeURIComponent(stripEndPrefix(firstEv)) + "&select=product_name,qty_paid&limit=5", "GET");
      Logger.log("同步後 Supabase 有 " + (afterSync ? afterSync.length : 0) + " 筆");
      if (afterSync && afterSync.length > 0) {
        afterSync.forEach(function(r){ Logger.log("  · " + r.product_name + " paid=" + r.qty_paid); });
      }
    }
  }
}


// =================================================================
// sendTWCompletedEmail：TW 已取貨感謝 email（自動，由 checkTakenEmails 觸發）
// =================================================================
function sendTWCompletedEmail(ss, orderId) {
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return;
  var rows = sheet.getDataRange().getValues();
  var tr = -1;
  for (var r = 1; r < rows.length; r++) {
    if (rows[r][0].toString() === orderId.toString()) { tr = r; break; }
  }
  if (tr === -1) return;

  var rowData   = rows[tr];
  var eventName = rowData[2] ? rowData[2].toString().trim() : "代購活動";
  var custEmail = rowData[10] ? rowData[10].toString().trim() : "";
  var custName  = rowData[8]  ? rowData[8].toString().trim()  : "客人";
  var summary   = rowData[3]  ? rowData[3].toString()         : "";
  if (!custEmail || custEmail.indexOf("@") === -1) return;

  var prog = "<div style='margin:28px 0;padding:0 10px;'><div style='position:relative;display:flex;justify-content:space-between;align-items:flex-start;'><div style='position:absolute;top:5px;left:0;right:0;height:1px;background:#e5e5e5;z-index:1;'></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#2980b9;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #2980b9;'></div><div style='font-size:11px;font-weight:700;color:#2980b9;margin-top:7px;'>已收到訂單</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#2980b9;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #2980b9;'></div><div style='font-size:11px;font-weight:700;color:#2980b9;margin-top:7px;'>已付款</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#2980b9;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #2980b9;'></div><div style='font-size:11px;font-weight:700;color:#2980b9;margin-top:7px;'>已到台</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#10b981;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #10b981;'></div><div style='font-size:11px;font-weight:700;color:#10b981;margin-top:7px;'>✓ 已完成取貨</div></div>" +
    "</div></div>";

  var socialBlock = "<div style='margin-top:14px;padding:14px;background:#f5f5f5;border-radius:8px;text-align:center;'>" +
    "<div style='font-size:11px;color:#999;margin-bottom:10px;'>追蹤我們，獲得最新代購資訊 👀</div>" +
    "<a href='https://www.threads.com/@886tw.81jp?igshid=NTc4MTIwNjQ2YQ==' style='display:inline-block;background:#111;color:#fff;font-size:12px;font-weight:700;padding:7px 16px;border-radius:16px;text-decoration:none;margin:3px;'>Threads @886tw.81jp</a>" +
    "<a href='https://www.instagram.com/886tw.81jp?igsh=MW8zMmVncGVwNmd1dg%3D%3D&utm_source=qr' style='display:inline-block;background:linear-gradient(135deg,#f09433,#e6683c,#dc2743,#cc2366,#bc1888);color:#fff;font-size:12px;font-weight:700;padding:7px 16px;border-radius:16px;text-decoration:none;margin:3px;'>Instagram @886tw.81jp</a>" +
    "</div>";

  var subject = "🎉 [已取貨] 代購 " + eventName + " — 886tw.81jp 台灣";
  var body = "<div style='font-family:Helvetica Neue,Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;color:#333;line-height:1.6;'>" +
    "<div style='background:linear-gradient(135deg,#2980b9,#1a5f8a);color:#fff;text-align:center;padding:28px 20px;'>" +
      "<div style='font-size:40px;margin-bottom:8px;'>🎁</div>" +
      "<div style='font-size:20px;font-weight:700;letter-spacing:1px;'>感謝您的支持！</div>" +
      "<div style='font-size:13px;opacity:.8;margin-top:6px;'>代購 " + eventName + "</div>" +
    "</div>" +
    "<div style='padding:28px 24px;'>" +
      "<p>您好，" + custName + "！</p>" +
      "<p style='margin-top:10px;font-size:14px;color:#555;'>您的訂單已完成取貨，感謝您選擇 <strong>886tw.81jp</strong> 的代購服務 🙏</p>" +
      "<p style='font-size:14px;background:#f2f2f2;padding:7px 12px;border-radius:4px;display:inline-block;margin:10px 0;'><strong>訂單編號：</strong>" + orderId + "</p>" +
      prog +
      "<div style='background:linear-gradient(135deg,#e3f2fd,#f0f7ff);border:1px solid #90caf9;border-radius:12px;padding:20px 24px;margin:20px 0;text-align:center;'>" +
        "<div style='font-size:36px;margin-bottom:8px;'>🎵</div>" +
        "<div style='font-size:15px;font-weight:600;color:#111;line-height:1.6;'>希望您喜歡這次的代購商品！<br>期待下次再為您服務 ✨</div>" +
        "<div style='font-size:13px;color:#777;margin-top:8px;'>如商品有任何問題，請於取貨後 <strong>7天內</strong> 聯繫我們</div>" +
      "</div>" +
      "<div style='margin-top:28px;padding-top:16px;border-top:1px dashed #ddd;font-size:12px;color:#777;line-height:1.9;'>" +
        "<p>如有任何問題，歡迎直接回覆此 Email 或於 Threads / Instagram 私訊我們查詢，謝謝！</p>" +
        socialBlock +
        "<p style='text-align:center;margin-top:10px;font-size:13px;color:#999;'><strong style='color:#333;'>886tw.81jp</strong></p>" +
      "</div>" +
    "</div></div>";

  MailApp.sendEmail({ to: custEmail, subject: subject, body: "感謝您完成取貨！訂單編號：" + orderId, htmlBody: body });
  Logger.log("[sendTWCompletedEmail] 已發送給 " + custEmail);
}

// =================================================================
// sendHKPickedUpEmail：HK 已取貨感謝 email（手動 / 批量）
// =================================================================
function sendHKPickedUpEmail(ss, orderId) {
  var sheet = ss.getSheetByName("訂單紀錄");
  if (!sheet) return { result: "error", message: "找不到訂單紀錄" };
  var rows = sheet.getDataRange().getValues();
  var tr = -1;
  for (var r = 1; r < rows.length; r++) {
    if (rows[r][0].toString() === orderId.toString()) { tr = r; break; }
  }
  if (tr === -1) return { result: "error", message: "找不到訂單" };

  var rowData   = rows[tr];
  var eventName = rowData[2] ? rowData[2].toString().trim() : "代購活動";
  var custEmail = rowData[10] ? rowData[10].toString().trim() : "";
  var custName  = rowData[8]  ? rowData[8].toString().trim()  : "客人";
  if (!custEmail || custEmail.indexOf("@") === -1) return { result: "error", message: "無效 Email" };

  var prog = "<div style='margin:28px 0;padding:0 10px;'><div style='position:relative;display:flex;justify-content:space-between;align-items:flex-start;'><div style='position:absolute;top:5px;left:0;right:0;height:1px;background:#e5e5e5;z-index:1;'></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>已收到訂單</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>已付款</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>到港途中</div></div>" +
    "<div style='position:relative;z-index:2;text-align:center;width:25%;'><div style='width:11px;height:11px;background:#111;border-radius:50%;margin:0 auto;border:2px solid #fff;box-shadow:0 0 0 1px #111;'></div><div style='font-size:11px;font-weight:700;color:#111;margin-top:7px;'>✓ 已寄出</div></div>" +
    "</div></div>";

  var socialBlock = "<div style='margin-top:14px;padding:14px;background:#f5f5f5;border-radius:8px;text-align:center;'>" +
    "<div style='font-size:11px;color:#999;margin-bottom:10px;'>追蹤我哋，獲得最新代購資訊 👀</div>" +
    "<a href='https://www.threads.net/@852hk.81jp' style='display:inline-block;background:#111;color:#fff;font-size:12px;font-weight:700;padding:7px 16px;border-radius:16px;text-decoration:none;'>Threads @852hk.81jp</a>" +
    "</div>";

  var subject = "🎉 [已取貨] 代購 " + eventName + " - 852hk.81jp";
  var body = "<div style='font-family:Helvetica Neue,Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;color:#333;line-height:1.6;'>" +
    "<div style='background:linear-gradient(135deg,#111,#333);color:#fff;text-align:center;padding:28px 20px;'>" +
      "<div style='font-size:40px;margin-bottom:8px;'>🎁</div>" +
      "<div style='font-size:20px;font-weight:700;letter-spacing:1px;'>多謝您的支持！</div>" +
      "<div style='font-size:13px;opacity:.8;margin-top:6px;'>代購 " + eventName + "</div>" +
    "</div>" +
    "<div style='padding:28px 24px;'>" +
      "<p>你好，" + custName + "！</p>" +
      "<p style='margin-top:10px;font-size:14px;color:#555;'>你嘅訂單已完成取件，多謝選擇 <strong>852hk.81jp</strong> 嘅代購服務 🙏</p>" +
      "<p style='font-size:14px;background:#f2f2f2;padding:7px 12px;border-radius:4px;display:inline-block;margin:10px 0;'><strong>訂單編號：</strong>" + orderId + "</p>" +
      prog +
      "<div style='background:#f5f5f5;border:1px solid #ddd;border-radius:12px;padding:20px 24px;margin:20px 0;text-align:center;'>" +
        "<div style='font-size:36px;margin-bottom:8px;'>🎵</div>" +
        "<div style='font-size:15px;font-weight:600;color:#111;line-height:1.6;'>希望您鐘意今次嘅代購商品！<br>期待下次再為您服務 ✨</div>" +
        "<div style='font-size:13px;color:#777;margin-top:8px;'>如商品有任何問題，請於取件後 <strong>7天內</strong> 聯繫我哋</div>" +
      "</div>" +
      "<div style='margin-top:28px;padding-top:16px;border-top:1px dashed #ddd;font-size:12px;color:#777;line-height:1.9;'>" +
        "<p>如有任何問題，歡迎直接回覆此 Email 或到 Threads inbox 我哋查詢，謝謝！</p>" +
        socialBlock +
        "<p style='text-align:center;margin-top:10px;font-size:13px;color:#999;'><strong style='color:#333;'>852hk.81jp</strong></p>" +
      "</div>" +
    "</div></div>";

  MailApp.sendEmail({ to: custEmail, subject: subject, body: "多謝您完成取件！訂單編號：" + orderId, htmlBody: body });
  sheet.getRange(tr + 1, 16).setValue("已完結");
  return { result: "success" };
}

// doPost handler for HK picked up (batch)
// =================================================================
// syncAllPurchaseRecords：一次過同步所有活動的買貨紀錄到 Supabase
// 第一次設定完 Supabase 後執行一次即可
// =================================================================
function syncAllPurchaseRecords() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var EXCL = ["訂單紀錄","易寄取地址","Blank","收單截止時間","購貨紀錄","盈利紀錄","[Data]盈利紀錄"];
  var sheets = ss.getSheets().map(function(s){ return s.getName(); })
    .filter(function(n){ return !n.startsWith("[Data]") && EXCL.indexOf(n) === -1; });

  var synced = 0;
  sheets.forEach(function(evName) {
    var pSheet = getEventPurchaseSheet(ss, evName, false);
    if (!pSheet) return;
    syncPurchaseRecordsToSupabase(evName);
    synced++;
    Logger.log("[syncAllPurchaseRecords] 完成：" + evName);
  });
  Logger.log("[syncAllPurchaseRecords] 共同步 " + synced + " 個活動");
  return synced;
}

// =================================================================
// onEdit：Sheet 改動時自動同步 Supabase
// 在 GAS 設定 onEdit trigger 指向此函式
// =================================================================
function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 易寄取地址改動 → 即時同步 HK addresses
    if (sheetName === "易寄取地址") {
      syncHKAddressesToSupabase();
      Logger.log("[onEdit] 易寄取地址 → Supabase 同步完成");
      return;
    }

    // 訂單紀錄 P欄（狀態）改動 → 不需要額外處理（付款/取消由各自函式處理）
    // 商品 sheet 改動 → 提示需手動同步（太頻繁不自動觸發）
  } catch(err) {
    Logger.log("[onEdit error] " + err.toString());
  }
}

// =================================================================
// setupSupabaseTriggers：設定 onEdit trigger（執行一次即可）
// =================================================================
function setupSupabaseTriggers() {
  // 移除舊的 onEdit triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "onEdit") ScriptApp.deleteTrigger(t);
  });
  // 重新建立
  ScriptApp.newTrigger("onEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  Logger.log("[setupSupabaseTriggers] onEdit trigger 設定完成");
}

// =================================================================
// 測試函式：手動在 GAS 執行，查看購貨紀錄 sheet 狀況
// 用法：選擇 debugPurchaseSheet，按執行，看執行記錄
// =================================================================
function debugPurchaseSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();

  // 列出所有 Data 開頭的 sheet
  Logger.log("=== 所有 [Data] 相關 Sheet ===");
  sheets.forEach(function(sh) {
    var n = sh.getName();
    if (n.indexOf("[Data]") === 0 || n.indexOf("購貨") !== -1) {
      var lastCol = sh.getLastColumn();
      var lastRow = sh.getLastRow();
      var row1 = sh.getRange(1, 1, 1, lastCol).getValues()[0];
      Logger.log("Sheet: " + n);
      Logger.log("  lastCol=" + lastCol + " lastRow=" + lastRow);
      Logger.log("  row1=" + JSON.stringify(row1));
      if (lastCol >= 8) {
        Logger.log("  ✅ H欄有資料：" + row1[7]);
      } else {
        Logger.log("  ❌ 只有 " + lastCol + " 欄，未有批次日期");
      }
    }
  });

  // 測試 getEventPurchaseSheet 能否找到特定活動
  // placeholder — do not remove
  Logger.log("\n=== 測試 getEventPurchaseSheet ===");
  // 列出所有非 Data / 非 system 的 sheet 名稱
  var EXCL = ["訂單紀錄","易寄取地址","Blank","收單截止時間","購貨紀錄","盈利紀錄"];
  var events = sheets.map(function(sh){ return sh.getName(); }).filter(function(n){
    return n.indexOf("[Data]") !== 0 && EXCL.indexOf(n) === -1;
  });
  Logger.log("活動列表: " + JSON.stringify(events));

  events.forEach(function(evName) {
    var clean = stripEndPrefix(evName);
    var targetName = "[Data](" + clean + ")購貨紀錄";
    var found = ss.getSheetByName(targetName);
    if (found) {
      var lc = found.getLastColumn();
      Logger.log("✅ " + evName + " → " + targetName + " (lastCol=" + lc + ")");
    } else {
      Logger.log("❌ " + evName + " → " + targetName + " 找不到!");
    }
  });
}

// =================================================================
// Nissen 代購 Email 模板
// =================================================================
function buildNissenEmailHtml_(order, isConfirmation) {
  var isHK = (order.region || "hk").toLowerCase() !== "tw";
  var currency = isHK ? "HK$" : "NT$";
  var items = order.items || [];
  var totalLocal = 0;

  var itemsHtml = "";
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var localPrice = isHK ? (it.price_hkd || it.local || 0) : (it.price_twd || it.local || 0);
    var qty = it.qty || 1;
    var subtotal = localPrice * qty;
    totalLocal += subtotal;
    var sub = [it.color, it.size, it.desc].filter(Boolean).join(" / ");
    var imgHtml = (it.image_url || it.image)
      ? "<img src=\"" + (it.image_url || it.image) + "\" style=\"width:48px;height:48px;object-fit:cover;border-radius:4px;margin-right:8px;vertical-align:middle;\">"
      : "";
    itemsHtml += "<tr>" +
      "<td style=\"padding:10px 12px;border-bottom:1px solid #f0f0f0;font-size:13px;\">" +
        imgHtml + "<span>" + it.name + (sub ? "<br><span style=\"font-size:11px;color:#888;\">" + sub + "</span>" : "") + "</span>" +
      "</td>" +
      "<td style=\"padding:10px 8px;border-bottom:1px solid #f0f0f0;font-size:13px;text-align:center;color:#555;\">" + qty + "</td>" +
      "<td style=\"padding:10px 12px;border-bottom:1px solid #f0f0f0;font-size:13px;text-align:right;\">" +
        (localPrice ? currency + Math.ceil(subtotal).toLocaleString() : "—") + "</td>" +
    "</tr>";
  }

  var payBlock = "";
  if (!isConfirmation) {
    // 通知 admin 的版本：簡潔顯示資料，admin 直接進 admin panel 處理
    payBlock = "<div style=\"background:#fffbeb;border-left:4px solid #f59e0b;padding:14px;border-radius:4px;margin-top:16px;font-size:13px;color:#92400e;\">" +
      "<p style=\"margin:0 0 6px;font-weight:700;\">🔔 新訂單通知</p>" +
      "<p style=\"margin:0;\">請到 <a href=\"" + ADMIN_PAGE_URL + "\" style=\"color:#d97706;\">Admin Panel</a> 查看訂單並跟進付款。</p>" +
    "</div>";
  } else {
    // 確認 email 給客人
    if (isHK) {
      payBlock = "<div style=\"background:#f0f7ff;border-left:4px solid #0066cc;padding:14px;border-radius:4px;margin-top:16px;font-size:13px;color:#1e3a5f;\">" +
        "<p style=\"margin:0 0 8px;font-weight:700;\">💰 付款方式（如未付款請盡快完成）</p>" +
        "<p style=\"margin:0 0 4px;\">PayMe：<a href=\"https://payme.hsbc/miru\" style=\"color:#ef4444;font-weight:600;\">payme.hsbc/miru</a></p>" +
        "<p style=\"margin:0;\">轉數快 (FPS)：<strong>8890873</strong>（Chow W. Y.）</p>" +
        "<p style=\"margin:8px 0 0;font-size:12px;color:#555;\">⚠️ 付款時請於備注填寫訂單號：<strong>" + order.order_no + "</strong></p>" +
      "</div>";
    } else {
      payBlock = "<div style=\"background:#f0fff4;border-left:4px solid #38a169;padding:14px;border-radius:4px;margin-top:16px;font-size:13px;color:#276749;\">" +
        "<p style=\"margin:0 0 8px;font-weight:700;\">💰 付款方式（郵局銀行匯款）</p>" +
        "<p style=\"margin:0 0 4px;\">銀行：<strong>中華郵政（郵局）</strong></p>" +
        "<p style=\"margin:0 0 4px;\">局號／帳號：<strong>0041860-0025565</strong></p>" +
        "<p style=\"margin:0;\">戶名：<strong>周◯恩</strong></p>" +
        "<p style=\"margin:8px 0 0;font-size:12px;color:#555;\">⚠️ 匯款時請於備註填寫訂單號後4碼：<strong>" + (order.order_no || "").slice(-4) + "</strong></p>" +
      "</div>";
    }
  }

  var greeting = isConfirmation
    ? "<p style=\"font-size:14px;\">您好 " + (order.customer_name || "") + "，</p><p style=\"font-size:14px;margin-top:6px;\">感謝您的訂購！" + (isHK ? "我哋" : "我們") + "已收到您的代購申請，以下是訂單確認資料。</p>"
    : "<p style=\"font-size:14px;\">收到新 Nissen 代購訂單，客戶資料如下：</p>";

  return "<div style=\"font-family:Helvetica Neue,Helvetica,Arial,sans-serif;max-width:580px;margin:0 auto;color:#333;line-height:1.6;padding:20px;background:#fafafa;\">" +
    "<h2 style=\"font-size:17px;font-weight:600;border-bottom:1px solid #e5e5e5;padding-bottom:12px;color:#111;\">" +
      (isConfirmation ? "✅ " : "📥 ") + "Nissen 代購訂單 — " + order.order_no +
    "</h2>" +
    greeting +
    "<div style=\"background:#fff;border:1px solid #e8e8e8;border-radius:8px;padding:14px;margin:14px 0;font-size:13px;line-height:1.8;\">" +
      "<strong>姓名：</strong>" + (order.customer_name || "—") + "<br>" +
      "<strong>Email：</strong>" + (order.customer_email || "—") + "<br>" +
      "<strong>電話：</strong>" + (order.customer_phone || "—") + "<br>" +
      "<strong>地區：</strong>" + (isHK ? "香港" : "台灣") + "<br>" +
      (order.remark ? "<strong>備注：</strong>" + order.remark : "") +
    "</div>" +
    "<h3 style=\"font-size:14px;font-weight:700;margin:16px 0 10px;color:#444;\">📋 訂購清單</h3>" +
    "<table style=\"width:100%;border-collapse:collapse;background:#fff;border:1px solid #eee;border-radius:6px;overflow:hidden;\">" +
      "<thead><tr style=\"background:#f7f7f7;\">" +
        "<th style=\"padding:8px 12px;text-align:left;font-size:12px;color:#666;\">商品</th>" +
        "<th style=\"padding:8px;text-align:center;font-size:12px;color:#666;width:50px;\">數量</th>" +
        "<th style=\"padding:8px 12px;text-align:right;font-size:12px;color:#666;width:80px;\">小計</th>" +
      "</tr></thead>" +
      "<tbody>" + itemsHtml + "</tbody>" +
    "</table>" +
    "<p style=\"text-align:right;font-size:15px;font-weight:700;margin-top:12px;\">" +
      "預計總金額：<span style=\"color:#c62828;\">" + (totalLocal ? currency + Math.ceil(totalLocal).toLocaleString() : "待確認") + "</span>" +
    "</p>" +
    "<p style=\"font-size:11px;color:#aaa;margin-top:0;\">※ 不含國際運費，最終以實際費用為準</p>" +
    payBlock +
    (isConfirmation
      ? "<p style=\"font-size:13px;margin-top:16px;\">📦 隨時查詢訂單進度：<a href=\"https://jpshopper-852hk81jp.vercel.app/status?id=" + encodeURIComponent(order.order_no) + "\" style=\"color:#2563eb;font-weight:600;\">按此查詢</a></p>"
      : "") +
    "<div style=\"margin-top:24px;padding-top:14px;border-top:1px dashed #ddd;font-size:12px;color:#999;\">" +
      (isConfirmation
        ? (isHK
          ? "<p>如有任何問題，歡迎直接回覆此 Email 或到 IG / Threads inbox 我哋查詢，謝謝！</p>"
          : "<p>如有任何問題，歡迎直接回覆此 Email 或 IG @886tw.81jp 私訊查詢，謝謝！</p>")
        : "<p>852hk.81jp Admin Panel 自動通知</p>") +
    "</div>" +
  "</div>";
}

