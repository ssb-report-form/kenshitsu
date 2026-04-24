// ═══════════════════════════════════════════════════════════════
// まいばすけっと検質報告書 - フロントエンド GAS連携コード
// index.html の <script> タグ前に読み込んでください
// <script src="kenshitsu_connector.js"></script>
// ═══════════════════════════════════════════════════════════════

// ── 設定（管理者が変更する箇所） ─────────────────────────────
const GAS_URL    = "https://script.google.com/macros/s/AKfycbx6E_AEWbV7jlhzqHZyeMJ9ZcVup9HkC09Gv1sxkdsRALieXnYLJdJPYONADqnT96g/exec";
const DRIVE_FOLDERS = {
  "大田": "1BlL8tKR9suHZEQlL6hegppF4r6qsQekt",  // 01.検質報告書_MYB大田
  "市川": "1G0pgkAz65rPdXSJQPJOCZzvQ6bcWZ0EY",  // 02.検質報告書_MYB市川
  "浮島": "1DrQep7gKl6irlDSY5d_SouLL_Jfp2Ult",  // 03.検質報告書_MYB浮島
};

// ── デモモード（GAS未接続時に自動切替） ─────────────────────
let DEMO_MODE = false;


// ═══════════════════════════════════════════════════════════════
// GAS通信ベース
// ═══════════════════════════════════════════════════════════════
// Netlify Functions プロキシ経由、またはダイレクトfetch
// GAS 302リダイレクト対策: redirect:'follow' + no-cors フォールバック

const PROXY_URL = "/.netlify/functions/gas-proxy";

async function callGAS(action, params = {}) {
  const payload = { action, ...params };

  // 1) Netlify Functionsプロキシ経由を試行
  try {
    const proxyResp = await fetch(PROXY_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (proxyResp.ok) {
      const json = await proxyResp.json();
      DEMO_MODE = false;
      updateStatus(true);
      return json;
    }
  } catch (e) {
    console.log("Proxy unavailable, trying direct GAS...");
  }

  // 2) ダイレクトGAS（GET + URLパラメータ）
  try {
    const url = new URL(GAS_URL);
    url.searchParams.set("action", action);
    Object.keys(params).forEach(k => {
      const v = params[k];
      url.searchParams.set(k, typeof v === "object" ? JSON.stringify(v) : v);
    });

    const resp = await fetch(url.toString(), {
      method: "GET",
      redirect: "follow",
    });
    if (resp.ok) {
      const json = await resp.json();
      DEMO_MODE = false;
      updateStatus(true);
      return json;
    }
  } catch (e) {
    console.log("Direct GAS failed:", e.message);
  }

  // 3) JSONP フォールバック
  try {
    const result = await callGASJsonp(action, params);
    if (result) {
      DEMO_MODE = false;
      updateStatus(true);
      return result;
    }
  } catch (e) {
    console.log("JSONP failed:", e.message);
  }

  // 4) すべて失敗 → デモモード
  console.warn("All GAS connections failed. Switching to DEMO mode.");
  DEMO_MODE = true;
  updateStatus(false);
  return null;
}

// JSONP呼び出し（GAS 302リダイレクト対策）
function callGASJsonp(action, params) {
  return new Promise((resolve, reject) => {
    const cbName = "_gasCallback_" + Date.now();
    const timeout = setTimeout(() => {
      delete window[cbName];
      const el = document.getElementById(cbName);
      if (el) el.remove();
      reject(new Error("JSONP timeout"));
    }, 30000); // Excel変換に時間がかかるため30秒

    window[cbName] = function(data) {
      clearTimeout(timeout);
      delete window[cbName];
      const el = document.getElementById(cbName);
      if (el) el.remove();
      resolve(data);
    };

    const url = new URL(GAS_URL);
    url.searchParams.set("action", action);
    url.searchParams.set("callback", cbName);
    Object.keys(params).forEach(k => {
      const v = params[k];
      url.searchParams.set(k, typeof v === "object" ? JSON.stringify(v) : v);
    });

    const script = document.createElement("script");
    script.id = cbName;
    script.src = url.toString();
    script.onerror = () => {
      clearTimeout(timeout);
      delete window[cbName];
      script.remove();
      reject(new Error("JSONP script error"));
    };
    document.head.appendChild(script);
  });
}


// ═══════════════════════════════════════════════════════════════
// API: 認証
// ═══════════════════════════════════════════════════════════════

async function gasAuth(center, password) {
  if (DEMO_MODE) return demoAuth(center, password);

  const result = await callGAS("auth", { center, password });
  if (!result) return demoAuth(center, password);

  if (result.success && result.data.authenticated) {
    return { ok: true, message: result.data.message };
  }
  return { ok: false, message: result.error || "認証失敗" };
}

function demoAuth(center, password) {
  const passwords = { "大田": "maibasu1", "市川": "maibasu2", "浮島": "maibasu3" };
  if (password === passwords[center]) {
    return { ok: true, message: center + "センターにログインしました（デモ）" };
  }
  return { ok: false, message: "パスワードが違います" };
}


// ═══════════════════════════════════════════════════════════════
// API: スタッフマスター取得
// ═══════════════════════════════════════════════════════════════

async function gasGetStaff(center) {
  // 1) GAS経由を試行
  if (!DEMO_MODE) {
    const result = await callGAS("getStaff", { center });
    if (result && result.success && result.data.staff && result.data.staff.length > 0) {
      return result.data.staff;
    }
  }

  // 2) Google Visualization API で直接スプレッドシートから読み取り
  try {
    const staff = await fetchStaffFromSheet(center);
    if (staff.length > 0) return staff;
  } catch (e) {
    console.warn("Sheets direct read failed:", e.message);
  }

  // 3) デモデータにフォールバック
  return demoGetStaff(center);
}

// Google Visualization API でセンター別スプレッドシートから担当者を取得
// シート名: 担当者マスタ
// A列 = 担当者名（1行目ヘッダー、2行目以降がデータ）

const REPORT_SS_IDS = {
  "大田": "13oodzjY7pnYiZ3VLdi6bOFutpiayLUfKlg63VQN7AMY",
  "市川": "183EOuR0Hk9x--v15TaGWAxpTJVfk-D1z26JpHTqmTA8",
  "浮島": "1DmidYcHZ8b8IPof5NAQuqurksfpGMtKYoxVG25IdIHk",
};

function fetchStaffFromSheet(center) {
  const ssId = REPORT_SS_IDS[center];
  if (!ssId) return Promise.reject(new Error("センター未設定: " + center));
  const sheetName = "担当者マスタ";

  // JSONP方式（file://プロトコルでもCORS問題なし）
  return new Promise(function(resolve, reject) {
    var cbName = "_staffCb_" + Date.now();
    var timeout = setTimeout(function() {
      delete window[cbName];
      var el = document.getElementById(cbName);
      if (el) el.remove();
      reject(new Error("Staff JSONP timeout"));
    }, 8000);

    window[cbName] = function(data) {
      clearTimeout(timeout);
      delete window[cbName];
      var el = document.getElementById(cbName);
      if (el) el.remove();

      if (!data || data.status !== "ok") {
        reject(new Error("Query error"));
        return;
      }
      var rows = data.table.rows || [];
      var staff = [];
      for (var i = 0; i < rows.length; i++) {
        var cell = rows[i].c[0];
        if (cell && cell.v) {
          var name = String(cell.v).trim();
          if (name && name !== "担当者名" && name !== "担当者") staff.push(name);
        }
      }
      resolve(staff);
    };

    // google.visualization.Query.setResponse を上書き
    window["google"] = window["google"] || {};
    window["google"]["visualization"] = window["google"]["visualization"] || {};
    window["google"]["visualization"]["Query"] = window["google"]["visualization"]["Query"] || {};
    window["google"]["visualization"]["Query"]["setResponse"] = window[cbName];

    var url = "https://docs.google.com/spreadsheets/d/" + ssId
      + "/gviz/tq?tqx=out:json&sheet=" + encodeURIComponent(sheetName)
      + "&tq=" + encodeURIComponent("SELECT A WHERE A IS NOT NULL");

    var script = document.createElement("script");
    script.id = cbName;
    script.src = url;
    script.onerror = function() {
      clearTimeout(timeout);
      delete window[cbName];
      script.remove();
      reject(new Error("Staff script load error"));
    };
    document.head.appendChild(script);
  });
}

function demoGetStaff(center) {
  const demoStaff = {
    "大田":  ["田中太郎", "佐藤花子", "鈴木一郎", "高橋美咲"],
    "市川":  ["渡辺健一", "伊藤由美", "山本大輔", "中村あゆみ"],
    "浮島":  ["小林修", "加藤恵子", "吉田拓也", "山田さくら"],
  };
  return demoStaff[center] || [];
}


// ═══════════════════════════════════════════════════════════════
// 取引先マスタ読み込み（JSONP方式）
// ═══════════════════════════════════════════════════════════════

// 取引先マスタ読み込み（A列=正式名、B列=特殊取引先名）
function fetchSupplierMaster(center) {
  var ssId = REPORT_SS_IDS[center];
  if (!ssId) return Promise.resolve({ names: [], aliases: {} });

  return new Promise(function(resolve, reject) {
    var cbName = "_supplierCb_" + Date.now();
    var timeout = setTimeout(function() {
      delete window[cbName];
      var el = document.getElementById(cbName);
      if (el) el.remove();
      resolve({ names: [], aliases: {} });
    }, 8000);

    window[cbName] = function(data) {
      clearTimeout(timeout);
      delete window[cbName];
      var el = document.getElementById(cbName);
      if (el) el.remove();
      if (!data || data.status !== "ok") { resolve({ names: [], aliases: {} }); return; }
      var names = [];
      var aliases = {};
      var rows = data.table.rows || [];
      for (var i = 0; i < rows.length; i++) {
        var cellA = rows[i].c[0];
        var cellB = rows[i].c[1];
        var official = cellA && cellA.v ? String(cellA.v).trim() : "";
        var alias = cellB && cellB.v ? String(cellB.v).trim() : "";
        if (!official || official === "取引先名" || official === "取引先") continue;
        names.push(official);
        if (alias) {
          aliases[alias] = official;
          aliases[alias.toLowerCase()] = official;
          aliases[alias.toUpperCase()] = official;
        }
      }
      resolve({ names: names, aliases: aliases });
    };

    window.google = window.google || {};
    window.google.visualization = window.google.visualization || {};
    window.google.visualization.Query = window.google.visualization.Query || {};
    window.google.visualization.Query.setResponse = window[cbName];

    var url = "https://docs.google.com/spreadsheets/d/" + ssId
      + "/gviz/tq?tqx=out:json&sheet=" + encodeURIComponent("取引先マスタ")
      + "&tq=" + encodeURIComponent("SELECT A,B");
    var script = document.createElement("script");
    script.id = cbName;
    script.src = url;
    script.onerror = function() {
      clearTimeout(timeout);
      delete window[cbName];
      script.remove();
      resolve({ names: [], aliases: {} });
    };
    document.head.appendChild(script);
  });
}

// 取引先名マッチング（特殊名→正式名、部分一致）
function matchSupplier(rawName, master) {
  if (!rawName) return rawName;
  var aliases = master.aliases || {};
  var names = master.names || [];

  // 1) 特殊取引先名で一致（大文字小文字無視）
  if (aliases[rawName]) return aliases[rawName];
  if (aliases[rawName.toLowerCase()]) return aliases[rawName.toLowerCase()];

  if (names.length === 0) return rawName;
  var norm = rawName.replace(/[\s　]/g, "");

  // 2) 正式名で完全一致
  for (var i = 0; i < names.length; i++) {
    if (names[i] === rawName) return names[i];
  }

  // 3) 部分一致
  var bestMatch = null;
  var bestScore = 0;
  for (var i = 0; i < names.length; i++) {
    var m = names[i];
    var mn = m.replace(/[\s　]/g, "");
    if (mn.indexOf(norm) >= 0 || norm.indexOf(mn) >= 0) {
      if (m.length > bestScore) { bestMatch = m; bestScore = m.length; }
      continue;
    }
    var stripped = mn.replace(/株式会社|（株）|\(株\)|有限会社/g, "");
    var rawStripped = norm.replace(/株式会社|（株）|\(株\)|有限会社/g, "");
    if (stripped.indexOf(rawStripped) >= 0 || rawStripped.indexOf(stripped) >= 0) {
      if (m.length > bestScore) { bestMatch = m; bestScore = m.length; }
    }
  }
  return bestMatch || rawName;
}


// ═══════════════════════════════════════════════════════════════
// API: 原価売価指示書データ取込（キャッシュ対応）
// ═══════════════════════════════════════════════════════════════

// localStorageキャッシュキー
function masterCacheKey(center) {
  return "kenshitsu_master_" + center;
}

// キャッシュ読み込み
function loadMasterCache(center) {
  try {
    var raw = localStorage.getItem(masterCacheKey(center));
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

// キャッシュ保存
function saveMasterCache(center, data) {
  try {
    data.cachedAt = new Date().toISOString();
    localStorage.setItem(masterCacheKey(center), JSON.stringify(data));
  } catch (e) {
    console.warn("Cache save failed:", e);
  }
}

// 店着日の計算（検質日+1）
function getDeliveryDate(inspDateStr) {
  var d = new Date(inspDateStr);
  d.setDate(d.getDate() + 1);
  return d.toISOString().split("T")[0];
}

// キャッシュの有効性チェック（店着日が対象期間内か）
function isCacheValid(cache, inspDateStr) {
  if (!cache || !cache.dateRange || !cache.dateRange.from || !cache.dateRange.to) return false;
  var delivery = getDeliveryDate(inspDateStr);
  return delivery >= cache.dateRange.from && delivery <= cache.dateRange.to;
}

// ファイル一覧取得
async function gasListFiles(center) {
  var result = await callGAS("listFiles", { center });
  if (result && result.success) {
    return result.data.files || [];
  }
  return [];
}

// GAS経由でExcel取込（取引先マスタで名寄せ）
async function gasImportExcel(center, fileId) {
  // 取引先マスタを並行読み込み
  var supplierMasterPromise = fetchSupplierMaster(center);

  if (DEMO_MODE) return demoImportExcel(center);

  var params = { center: center };
  if (fileId) params.fileId = fileId;
  var result = await callGAS("importExcel", params);
  if (!result) return demoImportExcel(center);

  if (result.success) {
    var supplierMaster = await supplierMasterPromise;
    var items = (result.data.items || []).map(function(it) {
      return {
        name: it.name, origin: it.origin,
        code: it.code || "", unit: it.unit || "",
        supplier: matchSupplier(it.supplier, supplierMaster),
      };
    });
    var importData = {
      items:     items,
      fileName:  result.data.fileName || "",
      fileDate:  result.data.fileDate || "",
      dateRange: result.data.dateRange || null,
    };
    saveMasterCache(center, importData);
    return importData;
  }
  console.warn("importExcel error:", result.error);
  return demoImportExcel(center);
}

function demoImportExcel(center) {
  // デモ用: 今週水曜〜翌火曜の対象期間を生成
  var now = new Date();
  var dayOfWeek = now.getDay(); // 0=日
  var diffToWed = (dayOfWeek >= 3) ? (dayOfWeek - 3) : (dayOfWeek + 4);
  var wed = new Date(now);
  wed.setDate(now.getDate() - diffToWed);
  var tue = new Date(wed);
  tue.setDate(wed.getDate() + 6);

  var items = [
    { name: "キャベツ",       origin: "群馬",       code: "4901001", unit: "CS", supplier: "丸果" },
    { name: "トマト",         origin: "熊本",       code: "4901002", unit: "CS", supplier: "大果" },
    { name: "レタス",         origin: "長野",       code: "4901003", unit: "CS", supplier: "丸果" },
    { name: "きゅうり",       origin: "埼玉",       code: "4901004", unit: "袋", supplier: "大果" },
    { name: "にんじん",       origin: "北海道",     code: "4901005", unit: "kg", supplier: "ベジテック" },
    { name: "ブロッコリー",   origin: "愛知",       code: "4901006", unit: "CS", supplier: "丸果" },
    { name: "ほうれん草",     origin: "群馬",       code: "4901007", unit: "束", supplier: "大果" },
    { name: "バナナ",         origin: "フィリピン", code: "4901008", unit: "箱", supplier: "ドール" },
    { name: "りんご",         origin: "青森",       code: "4901009", unit: "CS", supplier: "丸果" },
    { name: "みかん",         origin: "和歌山",     code: "4901010", unit: "kg", supplier: "ベジテック" },
    { name: "じゃがいも",     origin: "北海道",     code: "4901011", unit: "kg", supplier: "大果" },
    { name: "たまねぎ",       origin: "北海道",     code: "4901012", unit: "kg", supplier: "ベジテック" },
  ];

  var importData = {
    items: items,
    fileName: "【" + center + "】原価売価指示書（デモ）.xlsx",
    fileDate: new Date().toISOString(),
    dateRange: {
      from: wed.toISOString().split("T")[0],
      to:   tue.toISOString().split("T")[0],
    },
  };
  // デモでもキャッシュ保存
  saveMasterCache(center, importData);
  return importData;
}


// ═══════════════════════════════════════════════════════════════
// 検査中シート連携API（新フロー）
// ═══════════════════════════════════════════════════════════════

// 検査中シートへ1商品送信（POST no-cors）
async function gasSaveInspection(center, date, staff, item) {
  console.log('[inspection] 送信', { center: center, product: item.name });
  var payload = {
    action: 'saveInspection',
    center: center, date: date, staff: staff,
    item: JSON.stringify(item),
  };
  try {
    await fetch(GAS_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
    });
    return { ok: true };
  } catch(e) {
    console.error('[inspection] POST failed:', e.message);
    return { ok: false, message: e.message };
  }
}

// 検査中シートから当日分を取得（JSONP方式でレスポンス取得）
function gasGetInspections(center, deliveryDate) {
  return new Promise(function(resolve) {
    var cbName = "_getInspCb_" + Date.now();
    var timeout = setTimeout(function() {
      delete window[cbName];
      var el = document.getElementById(cbName); if (el) el.remove();
      resolve({ items: [] });
    }, 12000);

    window[cbName] = function(data) {
      clearTimeout(timeout);
      delete window[cbName];
      var el = document.getElementById(cbName); if (el) el.remove();
      if (data && data.success) resolve({ items: data.data.items || [] });
      else resolve({ items: [] });
    };

    var url = GAS_URL + '?action=getInspections'
      + '&center=' + encodeURIComponent(center)
      + '&deliveryDate=' + encodeURIComponent(deliveryDate || '')
      + '&callback=' + cbName;

    var script = document.createElement('script');
    script.id = cbName;
    script.src = url;
    script.onerror = function() {
      clearTimeout(timeout);
      delete window[cbName];
      script.remove();
      resolve({ items: [] });
    };
    document.head.appendChild(script);
  });
}

// 検査中シートの1件削除
async function gasDeleteInspection(center, id) {
  try {
    await fetch(GAS_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'deleteInspection', center: center, id: id }),
    });
    return { ok: true };
  } catch(e) { return { ok: false }; }
}

// 検査中シートを全削除（日付指定）
async function gasClearInspections(center, deliveryDate) {
  try {
    await fetch(GAS_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ action: 'clearInspections', center: center, deliveryDate: deliveryDate || '' }),
    });
    return { ok: true };
  } catch(e) { return { ok: false }; }
}


// ═══════════════════════════════════════════════════════════════
// 不良マスタ読み込み（JSONP方式）
// ═══════════════════════════════════════════════════════════════

function fetchDefectMaster(center) {
  var ssId = REPORT_SS_IDS[center];
  if (!ssId) return Promise.resolve([]);

  return new Promise(function(resolve) {
    var cbName = "_defectCb_" + Date.now();
    var timeout = setTimeout(function() {
      delete window[cbName]; var el = document.getElementById(cbName); if (el) el.remove();
      resolve([]);
    }, 8000);

    window[cbName] = function(data) {
      clearTimeout(timeout); delete window[cbName];
      var el = document.getElementById(cbName); if (el) el.remove();
      if (!data || data.status !== "ok") { resolve([]); return; }
      var reasons = (data.table.rows || [])
        .map(function(r) { return r.c[0] && r.c[0].v ? String(r.c[0].v).trim() : ""; })
        .filter(function(n) { return n && n !== "不良理由" && n !== "理由"; });
      resolve(reasons);
    };

    window.google = window.google || {};
    window.google.visualization = window.google.visualization || {};
    window.google.visualization.Query = window.google.visualization.Query || {};
    window.google.visualization.Query.setResponse = window[cbName];

    var url = "https://docs.google.com/spreadsheets/d/" + ssId
      + "/gviz/tq?tqx=out:json&sheet=" + encodeURIComponent("不良マスタ")
      + "&tq=" + encodeURIComponent("SELECT A WHERE A IS NOT NULL");
    var script = document.createElement("script");
    script.id = cbName; script.src = url;
    script.onerror = function() { clearTimeout(timeout); delete window[cbName]; script.remove(); resolve([]); };
    document.head.appendChild(script);
  });
}


// ═══════════════════════════════════════════════════════════════
// 途中保存（localStorage）
// ═══════════════════════════════════════════════════════════════

function progressKey(center, date) {
  return "kenshitsu_progress_" + center + "_" + date;
}

function saveProgress(center, date, items) {
  try {
    localStorage.setItem(progressKey(center, date), JSON.stringify({
      items: items, savedAt: new Date().toISOString()
    }));
  } catch (e) { console.warn("saveProgress failed:", e); }
}

function loadProgress(center, date) {
  try {
    var raw = localStorage.getItem(progressKey(center, date));
    return raw ? JSON.parse(raw) : null;
  } catch (e) { return null; }
}

function clearProgress(center, date) {
  try { localStorage.removeItem(progressKey(center, date)); } catch (e) {}
}


// ═══════════════════════════════════════════════════════════════
// API: 検査結果保存
// ═══════════════════════════════════════════════════════════════

async function gasSaveReport(center, date, staff, items) {
  console.log('[saveReport] 送信開始', { center: center, date: date, staff: staff, count: items.length });

  // 画像データを除去（画像は別途gasSaveDefectImageで送信済み）
  var lightItems = items.map(function(it) {
    var copy = {};
    for (var k in it) {
      if (k === 'inspPhotos' || k === 'defectPhotos') continue;
      copy[k] = it[k];
    }
    return copy;
  });

  // POST no-cors で送信（DEMO_MODEでも届く）
  var payload = {
    action: 'saveReport',
    center: center, date: date, staff: staff,
    items: JSON.stringify(lightItems),
  };

  try {
    await fetch(GAS_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
    });
    console.log('[saveReport] POST sent (no-cors)');
    return { ok: true, fireAndForget: true };
  } catch(e) {
    console.error('[saveReport] POST failed:', e.message);
  }
  return { ok: false, message: 'POST failed' };
}

function demoSaveReport() {
  return { ok: true, count: 0, defects: 0, demo: true };
}


// ═══════════════════════════════════════════════════════════════
// API: 不良画像をDriveに保存
// ═══════════════════════════════════════════════════════════════

async function gasSaveDefectImage(center, deliveryDate, productName, supplier, imageData, index) {
  console.log('[defect] 送信開始', { center: center, date: deliveryDate, product: productName, supplier: supplier, idx: index, size: (imageData||'').length });
  // DEMO_MODEでもno-cors POSTは届くので送信する

  var payload = {
    action: 'saveDefectImage',
    center: center, deliveryDate: deliveryDate,
    productName: productName, supplier: supplier || '',
    imageData: imageData, index: String(index),
  };

  // no-cors mode: レスポンスは読めないがPOSTは届く
  try {
    await fetch(GAS_URL, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
    });
    console.log('[defect] POST sent (no-cors)');
    return { ok: true, fireAndForget: true };
  } catch(e) {
    console.error('[defect] POST failed:', e.message);
  }

  // フォールバック: GET/JSONP（画像切り捨て）
  var result = await callGAS("saveDefectImage", {
    center: center, deliveryDate: deliveryDate,
    productName: productName, supplier: supplier || '',
    imageData: imageData, index: String(index),
  });
  if (!result) return { ok: true, demo: true };
  if (result.success) return { ok: true, fileUrl: result.data.fileUrl };
  return { ok: false, message: result.error };
}


// ═══════════════════════════════════════════════════════════════
// API: PDF生成 → Drive保存
// ═══════════════════════════════════════════════════════════════

async function gasSavePdf(center, date, staff, items) {
  if (DEMO_MODE) return demoSavePdf(center, date);

  // 画像データ含めてPOSTで送信（GETだとURLサイズ制限）
  var payload = {
    action: "savePdf",
    center: center, date: date, staff: staff, memo: "",
    items: JSON.stringify(items),
  };
  try {
    var resp = await fetch(GAS_URL, {
      method: "POST",
      headers: { "Content-Type": "text/plain" },
      body: JSON.stringify(payload),
      redirect: "follow",
    });
    if (resp.ok) {
      var json = await resp.json();
      if (json.success) {
        return { ok: true, fileId: json.data.fileId, fileName: json.data.fileName, fileUrl: json.data.fileUrl };
      }
      return { ok: false, message: json.error };
    }
  } catch(e) {
    console.log("POST savePdf failed:", e.message);
  }

  // フォールバック: JSONP（画像なし）
  var result = await callGAS("savePdf", {
    center: center, date: date, staff: staff, memo: "",
    items: JSON.stringify(items.map(function(it) {
      var copy = {}; for (var k in it) copy[k] = it[k];
      delete copy.inspPhotos; delete copy.defectPhotos;
      return copy;
    })),
  });

  if (!result) return demoSavePdf(center, date);

  if (result.success) {
    return { ok: true, fileId: result.data.fileId, fileName: result.data.fileName, fileUrl: result.data.fileUrl };
  }
  return { ok: false, message: result.error };
}

function demoSavePdf(center, date) {
  var dateStr = (date || "").replace(/-/g, "");
  var name = "【" + (center || "") + "農産】検質報告書_" + dateStr + "店着.pdf";
  return { ok: true, fileId: "demo", fileName: name, fileUrl: "#", demo: true };
}


// ═══════════════════════════════════════════════════════════════
// API: 接続テスト (ping)
// ═══════════════════════════════════════════════════════════════

async function gasPing() {
  const result = await callGAS("ping");
  if (result && result.success) {
    DEMO_MODE = false;
    updateStatus(true);
    return true;
  }
  DEMO_MODE = true;
  updateStatus(false);
  return false;
}


// ═══════════════════════════════════════════════════════════════
// ステータス表示の更新
// ═══════════════════════════════════════════════════════════════

function updateStatus(connected) {
  const dot = document.getElementById("statusDot");
  const txt = document.getElementById("statusTxt");
  if (!dot || !txt) return;

  if (connected) {
    dot.style.background = "var(--grn)";
    txt.textContent = "LIVE";
  } else {
    dot.style.background = "var(--amb)";
    txt.textContent = "DEMO";
  }
}


// ═══════════════════════════════════════════════════════════════
// ユーティリティ
// ═══════════════════════════════════════════════════════════════

function formatDateCompact(d) {
  return d.getFullYear()
    + String(d.getMonth() + 1).padStart(2, "0")
    + String(d.getDate()).padStart(2, "0");
}
