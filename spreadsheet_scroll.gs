// ═══════════════════════════════════════════════════════════════
// スプレッドシートを開いた時、各シートの最終データ行に自動スクロール
// ═══════════════════════════════════════════════════════════════
// 【使い方】
// 1. 対象スプレッドシートを開く
// 2. 拡張機能 → Apps Script
// 3. このコードを貼り付け → 保存（フロッピーアイコン）
// 4. 一度スプレッドシートを閉じて再度開く → 権限を許可
// ═══════════════════════════════════════════════════════════════

// 対象シート名（この名前のシートのみ最終行に移動）
var TARGET_SHEETS = ['検質報告', '検質不良'];

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  TARGET_SHEETS.forEach(function(name) {
    try {
      var sheet = ss.getSheetByName(name);
      if (!sheet) return;
      var lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.setActiveRange(sheet.getRange(lastRow, 1));
      }
    } catch (e) {}
  });

  // 現在表示中が対象シートなら最終行にスクロール
  var activeSheet = ss.getActiveSheet();
  if (TARGET_SHEETS.indexOf(activeSheet.getName()) >= 0) {
    var lastRow = activeSheet.getLastRow();
    if (lastRow > 1) {
      activeSheet.setActiveRange(activeSheet.getRange(lastRow, 1));
    }
  }
}

// 手動実行用（メニューから呼び出せる）
function goToLastRow() {
  onOpen();
  SpreadsheetApp.getActiveSpreadsheet().toast('最終行に移動しました', '完了', 3);
}

// カスタムメニューを追加
function createMenu() {
  SpreadsheetApp.getUi()
    .createMenu('🔽 ジャンプ')
    .addItem('最終行に移動', 'goToLastRow')
    .addToUi();
}

// onOpenはメニュー作成も呼ぶ
function onOpen_withMenu() {
  createMenu();
  onOpen();
}
