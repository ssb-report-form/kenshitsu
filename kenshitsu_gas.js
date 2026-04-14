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
  // C=●, K=産地, M=品名, Q=商品コード, S=発注単位, Z=取引先
  EXCEL_COLS: {
    FLAG:     2,   // C列 "●"
    ORIGIN:   10,  // K列 産地
    PRODUCT:  12,  // M列 品名
    CODE:     16,  // Q列 商品コード
    UNIT:     18,  // S列 発注単位
    SUPPLIER: 25,  // Z列 取引先名
  },

  // ── Google Drive フォルダID（不良画像保存先） ──────────────
  // 後でフォルダURLを設定してください
  DEFECT_IMAGE_FOLDERS: {
    "大田": "ここに大田用フォルダIDを貼り付け",
    "市川": "ここに市川用フォルダIDを貼り付け",
    "浮島": "ここに浮島用フォルダIDを貼り付け",
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
      case "saveDefectImage": return handleSaveDefectImage(data);
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
    case "saveReport":
      var reportItems = e.parameter.items ? JSON.parse(e.parameter.items) : [];
      result = handleSaveReport({ center: center, date: e.parameter.date || "", staff: e.parameter.staff || "", items: reportItems });
      break;
    case "savePdf":
      var pdfItems = e.parameter.items ? JSON.parse(e.parameter.items) : [];
      result = handleSavePdf({ center: center, date: e.parameter.date || "", staff: e.parameter.staff || "", memo: "", items: pdfItems });
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
        code:     String(row[C.CODE]     || "").trim(),
        unit:     String(row[C.UNIT]     || "").trim(),
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
  const { center, date, staff, items } = data;

  // 店着日 = 検査日+1
  var inspDate = new Date(date);
  var deliveryDate = new Date(inspDate);
  deliveryDate.setDate(deliveryDate.getDate() + 1);
  var deliveryStr = Utilities.formatDate(deliveryDate, "Asia/Tokyo", "yyyy-MM-dd");

  // センター別スプレッドシートを開く
  const ssId = CONFIG.REPORT_SPREADSHEETS[center];
  if (!ssId || ssId.indexOf("ここに") >= 0) {
    return error("センター「" + center + "」のスプレッドシートIDが未設定です");
  }
  const ss = SpreadsheetApp.openById(ssId);

  // ── シート「検質報告」 ── 全品目の検査結果
  var reportSheet = getOrCreateSheet(ss, "検質報告", [
    "タイムスタンプ", "店着日", "担当者", "商品コード", "商品名", "産地",
    "取引先", "発注単位", "入荷数", "検質数", "不良数", "不良率",
    "結果", "不良理由", "コメント"
  ]);

  // ── シート「検質不良」 ── 不良品のみ
  var defectSheet = getOrCreateSheet(ss, "検質不良", [
    "日付(店着日)", "取引先名", "商品コード", "商品名", "産地",
    "発注単位", "対象数(検質数合計)", "入荷数", "検質数(10%)",
    "不良数", "不良理由"
  ]);

  var timestamp = new Date().toLocaleString("ja-JP");
  var parsedItems = (typeof items === "string") ? JSON.parse(items) : items;
  var totalInspQty = 0;
  parsedItems.forEach(function(it) { totalInspQty += (Number(it.inspQty) || 0); });

  // 各品目を記録
  parsedItems.forEach(function(item) {
    var arrivalQty = Number(item.arrivalQty) || 0;
    var inspQty = Number(item.inspQty) || 0;
    var defectQty = Number(item.defectQty) || 0;
    var defectRate = inspQty > 0 ? Math.round(defectQty / inspQty * 1000) / 10 : 0;
    var result = defectQty > 0 ? "不良" : "合格";
    var reason = item.defectReason || "";
    if (reason === "その他（手入力）" && item.defectReasonText) {
      reason = "その他: " + item.defectReasonText;
    }

    // 検質報告シート
    reportSheet.appendRow([
      timestamp, deliveryStr, staff,
      item.code || "", item.name || "", item.origin || "",
      item.supplier || "", item.unit || "",
      arrivalQty, inspQty, defectQty,
      defectRate > 0 ? defectRate + "%" : "",
      result, reason, item.comment || ""
    ]);

    // 検質不良シート（不良品のみ）
    if (defectQty > 0) {
      defectSheet.appendRow([
        deliveryStr, item.supplier || "",
        item.code || "", item.name || "", item.origin || "",
        item.unit || "", totalInspQty,
        arrivalQty, inspQty, defectQty, reason
      ]);
    }
  });

  return ok({
    saved: true,
    count: parsedItems.length,
    defects: parsedItems.filter(function(i) { return (Number(i.defectQty) || 0) > 0; }).length,
    center: center,
    deliveryDate: deliveryStr
  });
}


// ═══════════════════════════════════════════════════════════════
// 5b. 不良画像をDriveに保存
// ═══════════════════════════════════════════════════════════════

function handleSaveDefectImage(data) {
  var center = data.center;
  var deliveryDate = data.deliveryDate || "";
  var productName = data.productName || "unknown";
  var imageData = data.imageData || "";  // base64
  var index = data.index || 1;

  var folderId = CONFIG.DEFECT_IMAGE_FOLDERS[center];
  if (!folderId || folderId.indexOf("ここに") >= 0) {
    return error("不良画像フォルダが未設定です（" + center + "）");
  }

  try {
    var folder = DriveApp.getFolderById(folderId);
    // base64 → Blob
    var base64 = imageData.replace(/^data:image\/\w+;base64,/, "");
    var blob = Utilities.newBlob(Utilities.base64Decode(base64), "image/jpeg",
      deliveryDate + "_" + productName + "_" + index + ".jpg");
    var file = folder.createFile(blob);
    return ok({ fileId: file.getId(), fileName: file.getName(), fileUrl: file.getUrl() });
  } catch (e) {
    Logger.log("saveDefectImage error: " + e);
    return error("画像保存に失敗: " + e.toString());
  }
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
  var parsedItems = (typeof items === "string") ? JSON.parse(items) : items;
  var hasDefect = parsedItems.some(function(i) { return Number(i.defectQty) > 0; });
  var totalArrival = 0, totalInsp = 0;
  parsedItems.forEach(function(i) {
    totalArrival += (Number(i.arrivalQty) || 0);
    totalInsp += (Number(i.inspQty) || 0);
  });
  var inspRate = totalArrival > 0 ? (Math.round(totalInsp / totalArrival * 1000) / 10) : 0;
  var totalPages = Math.ceil(parsedItems.length / 4);
  if (totalPages < 1) totalPages = 1;

  var TEL = {"大田":"03-5735-5107","市川":"047-395-7681","浮島":"044-276-1880"};
  var FAX = {"大田":"03-5735-5108","市川":"047-395-7682","浮島":"044-276-1881"};
  var tel = TEL[center] || "";
  var fax = FAX[center] || "";

  // 品目一覧HTML（4列×4行）
  var itemListHtml = '<table style="width:100%;font-size:9px">';
  for (var row = 0; row < 4; row++) {
    itemListHtml += '<tr>';
    for (var col = 0; col < 4; col++) {
      var idx = col * 4 + row;
      itemListHtml += '<td style="padding:1px 3px;width:25%">';
      if (idx < parsedItems.length) itemListHtml += (idx+1) + '.' + escapeHtml(parsedItems[idx].name);
      itemListHtml += '</td>';
    }
    itemListHtml += '</tr>';
  }
  itemListHtml += '</table>';

  var css = '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:sans-serif;font-size:9px;color:#333}' +
    'table{border-collapse:collapse}' +
    '.page{width:210mm;height:297mm;padding:8mm 10mm 10mm;position:relative;page-break-after:always;overflow:hidden}' +
    '.page:last-child{page-break-after:auto}' +
    '@media print{@page{size:A4 portrait;margin:0}.page{page-break-after:always;height:297mm}}';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' + css + '</style></head><body>';

  for (var page = 0; page < totalPages; page++) {
    var pageItems = parsedItems.slice(page * 4, (page + 1) * 4);

    // ページ内に不良があるかで抜取有無を判定
    var pageHasDefect = pageItems.some(function(it) { return Number(it.defectQty) > 0; });
    var samplingVal = pageHasDefect ? "有" : "無";
    var samplingBg = pageHasDefect ? "background:#ffff00;" : "";

    html += '<div class="page">';

    // ── ページ番号
    html += '<div style="position:absolute;top:8mm;right:10mm;font-size:10px">' + (page+1) + '/' + totalPages + '</div>';

    // ── タイトル
    html += '<div style="text-align:center;font-size:17px;font-weight:bold;letter-spacing:2px;margin-bottom:3mm">【まいばすけっと検質報告書】</div>';

    // ── 宛先
    html += '<div style="font-size:11px;font-weight:bold;margin-bottom:3mm">まいばすけっと株式会社　御中</div>';

    // ── 品目一覧 + センター名
    html += '<table style="width:100%;border:1px solid #999;margin-bottom:3mm"><tr>';
    html += '<td style="padding:2mm 3mm;border-right:1px solid #999;vertical-align:top">';
    html += '<div style="font-size:9px;font-weight:bold;margin-bottom:1.5mm">今週の検質品目</div>';
    html += itemListHtml;
    html += '</td>';
    html += '<td style="padding:2mm 3mm;width:38mm;vertical-align:middle;text-align:center">';
    html += '<div style="font-weight:bold;font-size:11px">' + escapeHtml(center) + '農産センター</div>';
    html += '</td></tr></table>';

    // ── サマリーバー
    html += '<table style="width:100%;border:1px solid #999;margin-bottom:4mm"><tr>';
    html += '<td style="padding:1.5mm 3mm;border-right:1px solid #999;width:17%"><div style="font-size:8px;color:#666">検質日</div><div style="font-weight:bold;font-size:13px">' + date + '</div></td>';
    html += '<td style="padding:1.5mm 3mm;border-right:1px solid #999;width:15%"><div style="font-size:8px;color:#666">報告者</div><div style="font-weight:bold;font-size:13px">' + escapeHtml(staff) + '</div></td>';
    html += '<td style="padding:2.5mm 4mm;border-right:1px solid #999;text-align:center;width:11%;' + samplingBg + '"><div style="font-size:8px;color:#555">抜取有無</div><div style="font-weight:bold;font-size:13px">' + samplingVal + '</div></td>';
    html += '<td style="padding:2.5mm 4mm;border-right:1px solid #999;text-align:center;width:17%"><div style="font-size:8px;color:#666">検質数計</div><div style="font-weight:bold;font-size:11px">' + totalInsp + ' ps</div></td>';
    html += '<td style="padding:2.5mm 4mm;border-right:1px solid #999;text-align:center;width:17%"><div style="font-size:8px;color:#666">対象数計</div><div style="font-weight:bold;font-size:11px">' + totalArrival + ' ps</div></td>';
    html += '<td style="padding:2.5mm 4mm;text-align:center;width:14%"><div style="font-size:8px;color:#666">検質率</div><div style="font-weight:bold;font-size:11px">' + inspRate + '%</div></td>';
    html += '</tr></table>';

    // ── 品目カード 2×2
    html += '<table style="width:100%" cellspacing="4"><tr>';
    for (var ci = 0; ci < 4; ci++) {
      if (ci === 2) html += '</tr><tr>';
      if (ci >= pageItems.length) {
        html += '<td style="width:50%"></td>';
        continue;
      }
      var item = pageItems[ci];
      var dq = Number(item.defectQty) || 0;
      var iq = Number(item.inspQty) || 0;
      var aq = Number(item.arrivalQty) || 0;
      var rate = iq > 0 ? (Math.round(dq / iq * 1000) / 10) : 0;
      var isNG = dq > 0;
      var reason = item.defectReason || "";
      if (reason === "その他（手入力）" && item.defectReasonText) reason = "その他: " + item.defectReasonText;
      var hBg = isNG ? "#c0392b" : "#1a5c2e";

      html += '<td style="width:50%;vertical-align:top;border:1.5px solid ' + hBg + ';padding:0">';

      // ヘッダー（商品名大きく）
      html += '<div style="background:' + hBg + ';color:#fff;padding:1.5mm 3mm;font-size:10px;font-weight:bold">';
      html += escapeHtml(item.name);
      if (isNG) html += '　<span style="font-size:8px">⚠ 不良あり</span>';
      html += '</div>';

      // ボディ
      html += '<div style="padding:1.5mm 3mm;font-size:9px">';

      // 仕入先(取引先名) / 産地
      html += '<table style="width:100%;margin-bottom:1.5mm"><tr>';
      html += '<td style="width:50%"><span style="color:#666;font-size:8px">仕入先</span><br><b>' + escapeHtml(item.supplier) + '</b></td>';
      html += '<td style="width:50%"><span style="color:#666;font-size:8px">産地</span><br><b>' + escapeHtml(item.origin) + '</b></td>';
      html += '</tr></table>';

      // 入荷数 / 検質数
      html += '<table style="width:100%;margin-bottom:1.5mm"><tr>';
      html += '<td style="width:50%"><span style="color:#666;font-size:8px">入荷数</span><br><b>' + aq + ' ps</b></td>';
      html += '<td style="width:50%"><span style="color:#666;font-size:8px">検質数</span><br><b>' + iq + ' ps</b></td>';
      html += '</tr></table>';

      // 不良数 / 不良率
      html += '<table style="width:100%;margin-bottom:1.5mm"><tr>';
      html += '<td style="width:50%"><span style="color:#666;font-size:8px">不良数</span><br><b style="color:' + (isNG ? '#c0392b' : '#333') + '">' + dq + ' ps</b></td>';
      html += '<td style="width:50%"><span style="color:#666;font-size:8px">不良率</span><br><b style="color:' + (isNG ? '#c0392b' : '#333') + '">' + rate + '%</b></td>';
      html += '</tr></table>';

      // 不良理由
      if (isNG && reason) {
        html += '<div style="padding:1mm 2mm;background:#fff5f5;border-radius:2px;border-left:3px solid #c0392b;font-size:8px;margin-bottom:1.5mm"><b>不良理由:</b> ' + escapeHtml(reason) + '</div>';
      }

      // コメント
      html += '<div style="font-size:8px;color:#555;margin-bottom:2mm;' + (isNG ? 'background:#fff5f5;padding:1mm 2mm;border-radius:2px' : '') + '">コメント: ' + escapeHtml(item.comment || '特に問題無し') + '</div>';

      // 検質画像（大きく横並び）
      var ip = item.inspPhotos || [];
      html += '<table style="width:100%"><tr>';
      html += buildPhotoCell(ip[0], '検質1', '#ccc', '#f5f5f5');
      html += '<td style="width:2mm"></td>';
      html += buildPhotoCell(ip[1], '検質2', '#ccc', '#f5f5f5');
      html += '</tr></table>';

      // 不良画像（不良時のみ）
      if (isNG) {
        var dp = item.defectPhotos || [];
        html += '<table style="width:100%;margin-top:2mm"><tr>';
        html += buildPhotoCell(dp[0], '不良1', '#c0392b', '#fff0f0');
        html += '<td style="width:2mm"></td>';
        html += buildPhotoCell(dp[1], '不良2', '#c0392b', '#fff0f0');
        html += '</tr></table>';
      }

      html += '</div></td>';
    }
    html += '</tr></table>';

    // フッター
    html += '<div style="position:absolute;bottom:8mm;left:10mm;right:10mm;border-top:1px solid #ccc;padding-top:2mm;font-size:8px;color:#666">';
    html += '<table style="width:100%"><tr>';
    html += '<td>' + escapeHtml(center) + '農産センター　TEL: ' + tel + '　FAX: ' + fax + '</td>';
    html += '<td style="text-align:right">' + (page+1) + ' / ' + totalPages + '</td>';
    html += '</tr></table></div>';

    html += '</div>'; // page
  }

  // ═══ 不良レポートページ（不良品がある場合のみ追加） ═══
  var defectItems = parsedItems.filter(function(it) { return Number(it.defectQty) > 0; });
  if (defectItems.length > 0) {
    html += '<div class="page">';
    html += '<div style="text-align:center;font-size:17px;font-weight:bold;letter-spacing:2px;margin-bottom:3mm">【まいばすけっと検質不良レポート】</div>';
    html += '<div style="font-size:11px;font-weight:bold;margin-bottom:4mm">まいばすけっと株式会社　御中</div>';

    // ヘッダー情報
    html += '<table style="width:100%;border:1px solid #999;margin-bottom:5mm"><tr>';
    html += '<td style="padding:2mm 4mm;border-right:1px solid #999;width:33%"><div style="font-size:8px;color:#666">検質日</div><div style="font-weight:bold;font-size:13px">' + date + '</div></td>';
    html += '<td style="padding:2mm 4mm;border-right:1px solid #999;width:34%"><div style="font-size:8px;color:#666">センター</div><div style="font-weight:bold;font-size:13px">' + escapeHtml(center) + '農産センター</div></td>';
    html += '<td style="padding:2mm 4mm;width:33%"><div style="font-size:8px;color:#666">報告者</div><div style="font-weight:bold;font-size:13px">' + escapeHtml(staff) + '</div></td>';
    html += '</tr></table>';

    // 不良品目
    defectItems.forEach(function(item) {
      var dq = Number(item.defectQty) || 0;
      var iq = Number(item.inspQty) || 0;
      var aq = Number(item.arrivalQty) || 0;
      var rate = iq > 0 ? (Math.round(dq / iq * 1000) / 10) : 0;
      var reason = item.defectReason || "";
      if (reason === "その他（手入力）" && item.defectReasonText) reason = "その他: " + item.defectReasonText;

      html += '<table style="width:100%;border:1.5px solid #c0392b;margin-bottom:4mm" cellspacing="0"><tr>';
      html += '<td colspan="2" style="background:#c0392b;color:#fff;padding:2mm 3mm;font-size:11px;font-weight:bold">⚠ ' + escapeHtml(item.name) + '</td>';
      html += '</tr><tr>';

      // 左：数値情報
      html += '<td style="padding:3mm;vertical-align:top;width:50%">';
      html += '<table style="width:100%;font-size:9px">';
      html += '<tr><td style="width:33%"><span style="color:#666;font-size:7.5px">仕入先</span><br><b>' + escapeHtml(item.supplier) + '</b></td>';
      html += '<td style="width:33%"><span style="color:#666;font-size:7.5px">産地</span><br><b>' + escapeHtml(item.origin) + '</b></td>';
      html += '<td style="width:34%"><span style="color:#666;font-size:7.5px">入荷数</span><br><b>' + aq + ' ps</b></td></tr>';
      html += '<tr><td><span style="color:#666;font-size:7.5px">検質数</span><br><b>' + iq + ' ps</b></td>';
      html += '<td><span style="color:#666;font-size:7.5px">不良数</span><br><b style="color:#c0392b">' + dq + ' ps</b></td>';
      html += '<td><span style="color:#666;font-size:7.5px">不良率</span><br><b style="color:#c0392b">' + rate + '%</b></td></tr>';
      html += '</table>';
      if (reason) html += '<div style="margin-top:2mm;padding:1.5mm 2mm;background:#fff5f5;border-radius:2px;font-size:8px"><b>不良理由:</b> ' + escapeHtml(reason) + '</div>';
      html += '<div style="margin-top:1.5mm;font-size:8px;color:#555">コメント: ' + escapeHtml(item.comment || '') + '</div>';
      html += '</td>';

      // 右：不良写真
      html += '<td style="padding:3mm;vertical-align:top;width:50%">';
      var dp = item.defectPhotos || [];
      html += '<table style="width:100%"><tr>';
      html += buildPhotoCell(dp[0], '不良写真1', '#c0392b', '#fff0f0');
      html += '<td style="width:2mm"></td>';
      html += buildPhotoCell(dp[1], '不良写真2', '#c0392b', '#fff0f0');
      html += '</tr></table>';
      html += '</td></tr></table>';
    });

    // フッター
    html += '<div style="position:absolute;bottom:8mm;left:10mm;right:10mm;border-top:1px solid #ccc;padding-top:2mm;font-size:8px;color:#666">';
    html += escapeHtml(center) + '農産センター　TEL: ' + tel + '　FAX: ' + fax;
    html += '</div>';
    html += '</div>'; // defect page
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

// 写真セル生成（base64があれば画像表示、なければプレースホルダー）
function buildPhotoCell(dataUrl, label, borderColor, bgColor) {
  if (dataUrl && dataUrl.indexOf('data:') === 0) {
    return '<td style="width:48%;height:100px;border:1px solid ' + borderColor + ';padding:2px;text-align:center;vertical-align:middle">'
      + '<img src="' + dataUrl + '" style="max-width:100%;max-height:96px;object-fit:contain">'
      + '</td>';
  }
  return '<td style="width:48%;height:100px;background:' + bgColor + ';border:1px solid ' + borderColor + ';text-align:center;vertical-align:middle;font-size:7px;color:#999">'
    + label + '</td>';
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
