// ═══════════════════════════════════════════════════════════════
// まいばすけっと検質報告書 - Google Apps Script（バックエンド）
// ═══════════════════════════════════════════════════════════════
// 【設定】ここだけ変更すれば動きます
// ═══════════════════════════════════════════════════════════════

const CONFIG = {

  // ── スプレッドシートID（共通：スタッフマスター等） ─────────
  SPREADSHEET_ID: "189ZSQAG1t2ld4zH_BEe434I7JbvScrPp8QPpRoLhL2o",

  // ── センター別スプレッドシートID（検査結果保存先） ─────────
  // 各センター専用のスプレッドシートIDを設定してください
  REPORT_SPREADSHEETS: {
    "大田": "13oodzjY7pnYiZ3VLdi6bOFutpiayLUfKlg63VQN7AMY",
    "市川": "183EOuR0Hk9x--v15TaGWAxpTJVfk-D1z26JpHTqmTA8",
    "浮島": "1DmidYcHZ8b8IPof5NAQuqurksfpGMtKYoxVG25IdIHk",
  },

  // ── Google Drive フォルダID（PDF保存先：センター別） ────────
  DRIVE_FOLDERS: {
    "大田": "1BlL8tKR9suHZEQlL6hegppF4r6qsQekt",  // 01.検質報告書_MYB大田
    "市川": "1G0pgkAz65rPdXSJQPJOCZzvQ6bcWZ0EY",  // 02.検質報告書_MYB市川
    "浮島": "1DrQep7gKl6irlDSY5d_SouLL_Jfp2Ult",  // 03.検質報告書_MYB浮島
  },

  // ── Google Drive フォルダID（原価売価指示書の取込元） ────────
  IMPORT_FOLDERS: {
    "大田": "1XCcBFwADDKHvLeuWe14sRB-xADyABN_F",
    "市川": "1FeeSAR2K0MdH1EZJRfsXrbA36X5fH4My",
    "浮島": "1YEdnp8LsF3AFHvo2C5lyqFQlArTc9mx3",
  },

  // ── 対象日セル位置（0-indexed） ───────────────────────────
  // AO=40, AP=41, AQ=42, AR=43, AS=44, AT=45, AU=46
  DATE_CELLS: {
    ROW: 3,          // 4行目（0-indexed=3）
    START_COL: 40,   // AO列
    END_COL: 46,     // AU列
  },

  // ── センター別パスワード ───────────────────────────────────
  PASSWORDS: {
    "大田":   "maibasu1",
    "市川":   "maibasu2",
    "浮島":   "maibasu3",
  },

  // ── Excel列マッピング（0-indexed） ─────────────────────────
  // C列=●フラグ, K列=産地, M列=品名, Z列=取引先名
  EXCEL_COLS: {
    FLAG:     2,   // C列 "●"
    ORIGIN:   10,  // K列 産地
    PRODUCT:  12,  // M列 品名
    SUPPLIER: 25,  // Z列 取引先名
  },

  // ── PDF設定 ────────────────────────────────────────────────
  PDF: {
    PAGE_SIZE:    "A4",
    ORIENTATION:  "portrait",   // 縦
    COLUMNS:      2,            // 2列
    ITEMS_PER_PAGE: 4,          // 4品/ページ
    HEADER_OK:    "#2ec98a",    // 合格ヘッダー色（緑）
    HEADER_NG:    "#e05565",    // 不良ヘッダー色（赤）
  },

  // ── スタッフマスター用シート名 ──────────────────────────────
  // 各センターのスプレッドシート内に「担当者マスタ」シートを作成
  STAFF_SHEET: "担当者マスタ",
};


// ═══════════════════════════════════════════════════════════════
// エントリーポイント
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case "auth":         return handleAuth(data);
      case "getMaster":    return handleGetMaster(data);
      case "getStaff":     return handleGetStaff(data);
      case "importExcel":  return handleImportExcel(data);
      case "listFiles":    return handleListFiles(data);
      case "saveReport":   return handleSaveReport(data);
      case "savePdf":      return handleSavePdf(data);
      default:             return ok({ message: "unknown action: " + data.action });
    }
  } catch (err) {
    Logger.log("doPost error: " + err);
    return error(err.toString());
  }
}

function doGet(e) {
  const action = e.parameter.action || "";
  const center = e.parameter.center || "";
  const callback = e.parameter.callback || "";

  var result;
  switch (action) {
    case "auth":
      result = handleAuth({ center: center, password: e.parameter.password || "" });
      break;
    case "getMaster":
      result = handleGetMaster({ center: center });
      break;
    case "getStaff":
      result = handleGetStaff({ center: center });
      break;
    case "importExcel":
      result = handleImportExcel({ center: center, fileId: e.parameter.fileId || "" });
      break;
    case "listFiles":
      result = handleListFiles({ center: center });
      break;
    case "ping":
      result = ok({ status: "alive", timestamp: new Date().toISOString() });
      break;
    default:
      result = ok({ message: "まいばすけっと検質報告書 API - OK" });
  }

  // JSONP対応: callbackパラメータがあればJavaScriptで返す
  if (callback) {
    var json = result.getContent();
    return ContentService
      .createTextOutput(callback + "(" + json + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return result;
}


// ═══════════════════════════════════════════════════════════════
// 1. 認証（センター + パスワード）
// ═══════════════════════════════════════════════════════════════

function handleAuth(data) {
  const { center, password } = data;

  if (!center || !CONFIG.PASSWORDS[center]) {
    return error("無効なセンター名です: " + center);
  }

  if (password !== CONFIG.PASSWORDS[center]) {
    return error("パスワードが違います");
  }

  return ok({
    authenticated: true,
    center: center,
    message: center + "センターにログインしました"
  });
}


// ═══════════════════════════════════════════════════════════════
// 2. スタッフマスター取得
// ═══════════════════════════════════════════════════════════════
// シート名: staff_大田, staff_市川, staff_浮島
// getSheets()で全シート走査 → prefix一致（文字化け対策）

function handleGetStaff(data) {
  const { center } = data;

  // センター別スプレッドシートから「担当者マスタ」シートを読み取り
  const ssId = CONFIG.REPORT_SPREADSHEETS[center];
  if (!ssId) {
    return ok({ staff: [], message: "センター「" + center + "」のスプレッドシートが未設定です" });
  }
  const ss = SpreadsheetApp.openById(ssId);

  // getSheetByNameの文字化け対策: 全シートを走査して一致検索
  let staffSheet = null;
  const targetName = CONFIG.STAFF_SHEET;
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const name = allSheets[i].getName();
    if (name === targetName || name.indexOf("担当者") >= 0) {
      staffSheet = allSheets[i];
      break;
    }
  }

  if (!staffSheet) {
    return ok({ staff: [], message: "「" + targetName + "」シートが見つかりません" });
  }

  const rows = staffSheet.getDataRange().getValues();
  // 1行目はヘッダー、2行目以降がスタッフ名
  const staff = [];
  for (let r = 1; r < rows.length; r++) {
    const name = String(rows[r][0]).trim();
    if (name) staff.push(name);
  }

  return ok({ staff: staff, center: center });
}


// ═══════════════════════════════════════════════════════════════
// 3. マスターデータ取得（全シートからセンター別データ）
// ═══════════════════════════════════════════════════════════════

function handleGetMaster(data) {
  const { center } = data;
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // master_大田 等のシートを探す
  let masterSheet = null;
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const name = allSheets[i].getName();
    if (name.indexOf("master_") === 0 && name.indexOf(center) >= 0) {
      masterSheet = allSheets[i];
      break;
    }
  }

  if (!masterSheet) {
    return ok({ items: [], message: "マスターシートが見つかりません" });
  }

  const rows = masterSheet.getDataRange().getValues();
  const items = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    items.push({
      name:     String(row[0] || "").trim(),
      origin:   String(row[1] || "").trim(),
      spec:     String(row[2] || "").trim(),
      supplier: String(row[3] || "").trim(),
    });
  }

  return ok({ items: items, center: center });
}


// ═══════════════════════════════════════════════════════════════
// 4. Excel取込（DriveからExcel読み取り → 品目データ返却）
// ═══════════════════════════════════════════════════════════════
// Drive advanced serviceの "Drive is not defined" エラー回避のため
// UrlFetchApp で直接読み取る

// ═══════════════════════════════════════════════════════════════
// 4a. ファイル一覧取得（ドロップダウン用）
// ═══════════════════════════════════════════════════════════════

function handleListFiles(data) {
  const { center } = data;
  const folderId = CONFIG.IMPORT_FOLDERS[center];
  if (!folderId) {
    return error("センター「" + center + "」のフォルダが未設定です");
  }
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  var list = [];
  while (files.hasNext()) {
    var f = files.next();
    var fname = f.getName();
    if (fname.indexOf("原価売価指示書") < 0) continue;
    var weekMatch = fname.match(/(\d+)週/);
    var weekNum = weekMatch ? parseInt(weekMatch[1], 10) : 0;
    var yearMatch = fname.match(/(\d+)年度/);
    var yearNum = yearMatch ? parseInt(yearMatch[1], 10) : 0;
    list.push({ id: f.getId(), name: fname, week: weekNum, year: yearNum, sortKey: yearNum * 100 + weekNum });
  }
  // 週番号の降順（最新が先頭）
  list.sort(function(a, b) { return b.sortKey - a.sortKey; });
  return ok({ files: list, center: center });
}


// ═══════════════════════════════════════════════════════════════
// 4b. Excel取込（指定ファイル or 最新）
// ═══════════════════════════════════════════════════════════════

function handleImportExcel(data) {
  const { center } = data;
  var fileId = data.fileId || "";

  // 原価売価指示書フォルダから取込
  const folderId = CONFIG.IMPORT_FOLDERS[center];
  if (!folderId) {
    return error("センター「" + center + "」の原価売価指示書フォルダが未設定です");
  }
  const folder = DriveApp.getFolderById(folderId);

  // フォルダ内の最新「原価売価指示書」Excelを探す
  // ファイル名の週番号（例:「7週」）が最大のものを選択
  const files = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  let latestFile = null;
  let latestDate = new Date(0);
  let latestWeek = -1;

  while (files.hasNext()) {
    const file = files.next();
    const fname = file.getName();
    if (fname.indexOf("原価売価指示書") < 0) continue;

    // ファイル名から週番号を抽出（例: "26年度7週" → 7）
    var weekMatch = fname.match(/(\d+)週/);
    var weekNum = weekMatch ? parseInt(weekMatch[1], 10) : 0;

    // 年度も考慮（例: "26年度" → 26）
    var yearMatch = fname.match(/(\d+)年度/);
    var yearNum = yearMatch ? parseInt(yearMatch[1], 10) : 0;
    var sortKey = yearNum * 100 + weekNum;

    if (sortKey > latestWeek) {
      latestWeek = sortKey;
      latestFile = file;
      latestDate = file.getDateCreated();
    }
  }

  // 「原価売価指示書」がなければフォルダ内の最新Excel（更新日順）
  if (!latestFile) {
    const allFiles = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
    while (allFiles.hasNext()) {
      const f = allFiles.next();
      if (f.getDateCreated() > latestDate) {
        latestFile = f;
        latestDate = f.getDateCreated();
      }
    }
  }

  // fileIdが指定されている場合はそのファイルを直接使用
  if (fileId) {
    try {
      latestFile = DriveApp.getFileById(fileId);
      latestDate = latestFile.getDateCreated();
    } catch (e) {
      return error("指定されたファイルが見つかりません: " + fileId);
    }
  }

  if (!latestFile) {
    return error("原価売価指示書が見つかりません（フォルダ: " + folderId + "）");
  }

  // Excelをスプレッドシートに一時変換して読み取り
  const tempSS = convertExcelToSheet(latestFile);
  if (!tempSS) {
    return error("Excelの変換に失敗しました");
  }

  try {
    const sheet = tempSS.getSheets()[0];
    const rows = sheet.getDataRange().getValues();
    const C = CONFIG.EXCEL_COLS;
    const D = CONFIG.DATE_CELLS;

    // ── 対象日情報を読み取り（AO4〜AU4 = 日付、AO5〜AU5 = 曜日）
    var dateFrom = null;
    var dateTo = null;
    var dates = [];
    for (var col = D.START_COL; col <= D.END_COL; col++) {
      var cellVal = rows[D.ROW] ? rows[D.ROW][col] : null;
      if (cellVal instanceof Date) {
        dates.push(cellVal.toISOString().split("T")[0]);
        if (!dateFrom || cellVal < new Date(dateFrom)) dateFrom = cellVal.toISOString().split("T")[0];
        if (!dateTo || cellVal > new Date(dateTo)) dateTo = cellVal.toISOString().split("T")[0];
      } else if (cellVal) {
        // 文字列の日付もパース試行
        var parsed = new Date(cellVal);
        if (!isNaN(parsed.getTime())) {
          var ds = parsed.toISOString().split("T")[0];
          dates.push(ds);
          if (!dateFrom || parsed < new Date(dateFrom)) dateFrom = ds;
          if (!dateTo || parsed > new Date(dateTo)) dateTo = ds;
        }
      }
    }

    // ── 取引先マスタを読み込み（名寄せ用）
    var supplierMaster = loadSupplierMaster(center);

    // ── 品目データ読み取り（●フラグがある行）
    const items = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      var flag = String(row[C.FLAG] || "").trim();
      if (flag !== "●") continue;

      var productName = String(row[C.PRODUCT] || "").trim();
      if (!productName) continue;

      var rawSupplier = String(row[C.SUPPLIER] || "").trim();
      var matchedSupplier = matchSupplierName(rawSupplier, supplierMaster);

      items.push({
        name:     productName,
        origin:   String(row[C.ORIGIN]   || "").trim(),
        supplier: matchedSupplier,
      });
    }

    return ok({
      items:     items,
      fileName:  latestFile.getName(),
      fileDate:  latestDate.toISOString(),
      center:    center,
      dateRange: { from: dateFrom, to: dateTo, dates: dates },
    });

  } finally {
    // 一時ファイルを削除
    DriveApp.getFileById(tempSS.getId()).setTrashed(true);
  }
}

// ExcelをGoogleスプレッドシートに一時変換
// GASの「サービス」で Drive API を有効化してください
function convertExcelToSheet(excelFile) {
  try {
    var blob = excelFile.getBlob();
    // Drive Advanced Service (v3) で変換
    var resource = {
      name: "_temp_kenshitsu_" + new Date().getTime(),
      mimeType: "application/vnd.google-apps.spreadsheet"
    };
    var file = Drive.Files.create(resource, blob);
    return SpreadsheetApp.openById(file.id);
  } catch (e) {
    Logger.log("convertExcelToSheet error: " + e);
    return null;
  }
}


// ═══════════════════════════════════════════════════════════════
// 5. 検査結果保存（スプレッドシートに記録）
// ═══════════════════════════════════════════════════════════════

function handleSaveReport(data) {
  const { center, date, staff, memo, items } = data;

  // センター別スプレッドシートを開く
  const ssId = CONFIG.REPORT_SPREADSHEETS[center];
  if (!ssId || ssId.indexOf("ここに") >= 0) {
    return error("センター「" + center + "」の検査結果用スプレッドシートIDが未設定です");
  }
  const ss = SpreadsheetApp.openById(ssId);

  let reportSheet = getOrCreateSheet(ss, "reports", [
    "タイムスタンプ", "検査日", "担当者",
    "品目名", "産地", "規格", "仕入先", "結果", "不良種別", "不良詳細"
  ]);

  const timestamp = new Date().toLocaleString("ja-JP");

  // 各品目を1行ずつ記録
  items.forEach(function(item) {
    reportSheet.appendRow([
      timestamp,
      date,
      staff,
      item.name,
      item.origin,
      item.spec,
      item.supplier,
      item.result === "pass" ? "合格" : item.result === "fail" ? "不良" : "未検査",
      item.defectType || "",
      item.defectNote || "",
    ]);
  });

  return ok({
    saved: true,
    count: items.length,
    center: center,
    date: date
  });
}


// ═══════════════════════════════════════════════════════════════
// 6. PDF生成 → Driveに保存
// ═══════════════════════════════════════════════════════════════
// A4縦 / 2列 / 4品ページ / 緑(合格)・赤(不良)ヘッダー
// 不良品があれば別ページ追加

function handleSavePdf(data) {
  const { center, date, staff, memo, items } = data;
  const P = CONFIG.PDF;

  // HTMLテンプレートでPDF生成
  const html = buildPdfHtml(center, date, staff, memo, items);
  // タイトル例: 【大田農産】検質報告書_20260409店着.pdf
  const dateStr = date.replace(/-/g, "");
  const pdfName = "【" + center + "農産】検質報告書_" + dateStr + "店着.pdf";

  const blob = HtmlService.createHtmlOutput(html)
    .getBlob()
    .setName(pdfName);

  // センター別フォルダに保存
  const folderId = CONFIG.DRIVE_FOLDERS[center];
  if (!folderId) {
    return error("センター「" + center + "」の保存先フォルダが未設定です");
  }
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);

  return ok({
    saved: true,
    fileId: file.getId(),
    fileName: file.getName(),
    fileUrl: file.getUrl(),
  });
}

function buildPdfHtml(center, date, staff, memo, items) {
  const P = CONFIG.PDF;
  const pass = items.filter(function(i) { return i.result === "pass"; }).length;
  const fail = items.filter(function(i) { return i.result === "fail"; }).length;
  const defects = items.filter(function(i) { return i.result === "fail"; });
  const hasDefect = defects.length > 0;

  let html = '<!DOCTYPE html><html><head><meta charset="UTF-8">';
  html += '<style>';
  html += 'body{font-family:"Noto Sans JP",sans-serif;font-size:10px;margin:20px;color:#333}';
  html += '.header{text-align:center;padding:12px;border-bottom:3px solid ' + (hasDefect ? P.HEADER_NG : P.HEADER_OK) + ';margin-bottom:16px}';
  html += '.title{font-size:16px;font-weight:700}';
  html += '.info{font-size:10px;color:#666;margin-top:4px}';
  html += '.summary{display:flex;justify-content:space-between;margin-bottom:12px;font-size:11px}';
  html += '.grid{display:grid;grid-template-columns:1fr 1fr;gap:8px}';
  html += '.item{border:1px solid #ddd;border-radius:4px;padding:8px}';
  html += '.item-name{font-weight:700;margin-bottom:2px}';
  html += '.item-meta{color:#666;font-size:9px}';
  html += '.ok{color:' + P.HEADER_OK + ';font-weight:700}';
  html += '.ng{color:' + P.HEADER_NG + ';font-weight:700}';
  html += '.defect-page{page-break-before:always;margin-top:20px}';
  html += '.defect-header{text-align:center;padding:12px;border-bottom:3px solid ' + P.HEADER_NG + ';margin-bottom:16px}';
  html += '.defect-item{border:1px solid ' + P.HEADER_NG + ';border-radius:4px;padding:10px;margin-bottom:8px}';
  html += '</style></head><body>';

  // ── ヘッダー
  html += '<div class="header">';
  html += '<div class="title">まいばすけっと 検質報告書</div>';
  html += '<div class="info">' + center + 'センター ｜ ' + date + ' ｜ 担当: ' + staff + '</div>';
  if (memo) html += '<div class="info">備考: ' + memo + '</div>';
  html += '</div>';

  // ── サマリー
  html += '<div class="summary">';
  html += '<span>品目数: ' + items.length + '件</span>';
  html += '<span class="ok">合格: ' + pass + '</span>';
  html += '<span class="ng">不良: ' + fail + '</span>';
  html += '</div>';

  // ── 品目グリッド（2列、4品/ページ）
  html += '<div class="grid">';
  items.forEach(function(item, idx) {
    // ページ区切り (4品ごと、ただし最初は除く)
    if (idx > 0 && idx % P.ITEMS_PER_PAGE === 0) {
      html += '</div><div style="page-break-before:always"></div><div class="grid">';
    }
    var resultClass = item.result === "pass" ? "ok" : item.result === "fail" ? "ng" : "";
    var resultText  = item.result === "pass" ? "合格" : item.result === "fail" ? "不良" : "未検査";
    html += '<div class="item">';
    html += '<div class="item-name">' + (idx + 1) + '. ' + escapeHtml(item.name) + '</div>';
    html += '<div class="item-meta">' + escapeHtml(item.origin) + ' / ' + escapeHtml(item.spec) + ' / ' + escapeHtml(item.supplier) + '</div>';
    html += '<div class="' + resultClass + '">' + resultText + '</div>';
    html += '</div>';
  });
  html += '</div>';

  // ── 不良品レポートページ（不良品がある場合のみ追加）
  if (hasDefect) {
    html += '<div class="defect-page">';
    html += '<div class="defect-header"><div class="title">不良品レポート</div>';
    html += '<div class="info">' + center + 'センター ｜ ' + date + ' ｜ 不良 ' + fail + '件</div></div>';

    defects.forEach(function(d, idx) {
      html += '<div class="defect-item">';
      html += '<div style="font-weight:700;color:' + CONFIG.PDF.HEADER_NG + '">' + escapeHtml(d.name) + '</div>';
      html += '<div class="item-meta">' + escapeHtml(d.origin) + ' / ' + escapeHtml(d.spec) + ' / ' + escapeHtml(d.supplier) + '</div>';
      html += '<div style="margin-top:4px"><strong>種別:</strong> ' + escapeHtml(d.defectType || "未選択") + '</div>';
      html += '<div><strong>詳細:</strong> ' + escapeHtml(d.defectNote || "なし") + '</div>';
      html += '</div>';
    });

    html += '</div>';
  }

  html += '</body></html>';
  return html;
}


// ═══════════════════════════════════════════════════════════════
// ユーティリティ
// ═══════════════════════════════════════════════════════════════

// ── 取引先マスタ読み込み ─────────────────────────────────────
// センター別スプレッドシートの「取引先マスタ」シートからA列を読み取り
// 戻り値: { names: [正式名,...], aliases: { "KIFA": "株式会社ケーアイ・フレッシュアクセス", ... } }
function loadSupplierMaster(center) {
  var ssId = CONFIG.REPORT_SPREADSHEETS[center];
  if (!ssId) return { names: [], aliases: {} };
  try {
    var ss = SpreadsheetApp.openById(ssId);
    var allSheets = ss.getSheets();
    var sheet = null;
    for (var i = 0; i < allSheets.length; i++) {
      var name = allSheets[i].getName();
      if (name === "取引先マスタ" || name.indexOf("取引先") >= 0) {
        sheet = allSheets[i];
        break;
      }
    }
    if (!sheet) return { names: [], aliases: {} };
    var rows = sheet.getDataRange().getValues();
    var names = [];
    var aliases = {};
    for (var r = 1; r < rows.length; r++) {
      var official = String(rows[r][0] || "").trim();  // A列: 正式名
      var alias = String(rows[r][1] || "").trim();      // B列: 特殊取引先名
      if (official) {
        names.push(official);
        if (alias) {
          aliases[alias] = official;
          aliases[alias.toLowerCase()] = official;
          aliases[alias.toUpperCase()] = official;
        }
      }
    }
    return { names: names, aliases: aliases };
  } catch (e) {
    Logger.log("loadSupplierMaster error: " + e);
    return { names: [], aliases: {} };
  }
}

// ── 取引先名マッチング（省略名→正式名） ─────────────────────
// 部分一致で最も長くマッチする正式名を返す
function matchSupplierName(rawName, master) {
  if (!rawName) return rawName;
  var names = master.names || [];
  var aliases = master.aliases || {};

  // 1) 特殊取引先名（B列）で完全一致（大文字小文字無視）
  if (aliases[rawName]) return aliases[rawName];
  if (aliases[rawName.toLowerCase()]) return aliases[rawName.toLowerCase()];
  if (aliases[rawName.toUpperCase()]) return aliases[rawName.toUpperCase()];

  if (names.length === 0) return rawName;

  var norm = rawName.replace(/[\s　]/g, "");

  // 2) 正式名（A列）で完全一致
  for (var i = 0; i < names.length; i++) {
    if (names[i] === rawName) return names[i];
  }

  // 3) 部分一致
  var bestMatch = null;
  var bestScore = 0;
  for (var i = 0; i < names.length; i++) {
    var masterName = names[i];
    var masterNorm = masterName.replace(/[\s　]/g, "");

    if (masterNorm.indexOf(norm) >= 0 || norm.indexOf(masterNorm) >= 0) {
      if (masterName.length > bestScore) {
        bestMatch = masterName;
        bestScore = masterName.length;
      }
      continue;
    }

    var stripped = masterNorm.replace(/株式会社|（株）|\(株\)|有限会社/g, "");
    var rawStripped = norm.replace(/株式会社|（株）|\(株\)|有限会社/g, "");
    if (stripped.indexOf(rawStripped) >= 0 || rawStripped.indexOf(stripped) >= 0) {
      if (masterName.length > bestScore) {
        bestMatch = masterName;
        bestScore = masterName.length;
      }
    }
  }

  return bestMatch || rawName;
}

function getOrCreateSheet(ss, name, headers) {
  var sh = null;
  var allSheets = ss.getSheets();
  for (var i = 0; i < allSheets.length; i++) {
    if (allSheets[i].getName() === name) {
      sh = allSheets[i];
      break;
    }
  }
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    // ヘッダー行を色付け
    sh.getRange(1, 1, 1, headers.length)
      .setBackground("#1e2e4a")
      .setFontColor("#dde6f4")
      .setFontWeight("bold");
  }
  return sh;
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function error(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}
