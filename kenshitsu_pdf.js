// ═══════════════════════════════════════════════════════════════
// まいばすけっと検質報告書 - PDF生成（ブラウザ側）
// window.print() でPDF出力。GAS不要。
// ═══════════════════════════════════════════════════════════════

var CENTER_TEL = {"大田":"03-5735-5107","市川":"047-395-7681","浮島":"044-276-1880"};
var CENTER_FAX = {"大田":"03-5735-5108","市川":"047-395-7682","浮島":"044-276-1881"};

// HTML生成のみ（プレビュー用）
function generatePdfHtml(center, date, staff, sampling, items) {
  return _buildPdfFullHtml(center, date, staff, sampling, items, false);
}

// HTML生成＋印刷ウィンドウ
function generateAndPrintPdf(center, date, staff, sampling, items) {
  return _buildPdfFullHtml(center, date, staff, sampling, items, true);
}

function _buildPdfFullHtml(center, date, staff, sampling, items, doPrint) {
  var tel = CENTER_TEL[center] || "";
  var fax = CENTER_FAX[center] || "";
  var totalArrival = 0, totalInsp = 0;
  items.forEach(function(i) {
    totalArrival += (Number(i.arrivalQty) || 0);
    totalInsp += (Number(i.inspQty) || 0);
  });
  var inspRate = totalArrival > 0 ? (Math.round(totalInsp / totalArrival * 1000) / 10) : 0;
  var hasDefect = items.some(function(i) { return Number(i.defectQty) > 0; });

  // 品目一覧HTML（4列×4行）
  var itemListHtml = '';
  for (var row = 0; row < 4; row++) {
    for (var col = 0; col < 4; col++) {
      var idx = col * 4 + row;
      itemListHtml += '<div>' + (idx < items.length ? (idx+1) + '.' + esc(items[idx].name) : '') + '</div>';
    }
  }

  // ページ分割（4品/ページ）
  var pages = [];
  for (var i = 0; i < items.length; i += 4) pages.push(items.slice(i, i + 4));
  if (pages.length === 0) pages.push([]);
  var totalPages = pages.length;

  // 品目カードHTML生成
  function buildCard(item) {
    var dq = Number(item.defectQty) || 0;
    var iq = Number(item.inspQty) || 0;
    var aq = Number(item.arrivalQty) || 0;
    var rate = iq > 0 ? (Math.round(dq / iq * 1000) / 10) : 0;
    var isNG = dq > 0;
    var bc = isNG ? '#c0392b' : '#1a5c2e';
    var reason = item.defectReason || '';
    if (reason === 'その他（手入力）' && item.defectReasonText) reason = 'その他: ' + item.defectReasonText;

    // 検質写真（高さ固定で縦長画像もはみ出さない）
    var photoH = '48mm';
    var photosHtml = '';
    var ip = item.inspPhotos || [];
    for (var p = 0; p < 2; p++) {
      if (ip[p]) {
        photosHtml += '<div style="flex:1;height:' + photoH + ';border:1px solid #ccc;overflow:hidden;"><img src="' + ip[p] + '" style="width:100%;height:100%;object-fit:contain;display:block;background:#f9f9f9;"></div>';
      } else {
        photosHtml += '<div style="flex:1;height:' + photoH + ';border:1px dashed #ccc;display:flex;align-items:center;justify-content:center;font-size:8px;color:#aaa;">写真' + (p+1) + '</div>';
      }
    }

    var h = '<div style="border-radius:4px;overflow:hidden;border:1.5px solid ' + bc + ';display:flex;flex-direction:column;">';
    // ヘッダー
    h += '<div style="background:' + bc + ';color:#fff;padding:1.5mm 3mm;font-size:10px;font-weight:700;">' + esc(item.name);
    if (isNG) h += '　<span style="font-size:8px">⚠ 不良あり</span>';
    h += '</div>';
    // ボディ
    h += '<div style="padding:1.5mm 3mm;font-size:9px;flex:1;display:flex;flex-direction:column;">';
    h += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:1.5mm;margin-bottom:2mm;">';
    h += '<div><span style="color:#000;font-size:8px;">仕入先</span><br>' + esc(item.supplier || '-') + '</div>';
    h += '<div><span style="color:#000;font-size:8px;">産地</span><br>' + esc(item.origin || '-') + '</div>';
    h += '<div><span style="color:#000;font-size:8px;">入荷数</span><br><b>' + aq + ' ps</b></div>';
    h += '<div><span style="color:#000;font-size:8px;">検質数</span><br><b>' + iq + ' ps</b></div>';
    h += '<div><span style="color:#000;font-size:8px;">不良数</span><br><b' + (isNG ? ' style="color:#c0392b;"' : '') + '>' + dq + ' ps</b></div>';
    h += '<div><span style="color:#000;font-size:8px;">不良率</span><br><b' + (isNG ? ' style="color:#c0392b;"' : '') + '>' + rate + '%</b></div>';
    h += '</div>';
    if (isNG && reason) {
      h += '<div style="background:#fff5f5;border-radius:2px;padding:1.5mm 2mm;border-left:3px solid #c0392b;font-size:8px;margin-bottom:1.5mm;"><b>不良理由:</b> ' + esc(reason) + '</div>';
    }
    h += '<div style="border-top:1px solid #eee;padding-top:1.5mm;margin-bottom:2mm;font-size:9px;' + (isNG ? 'background:#fff5f5;border-radius:2px;padding:1.5mm 2mm;border-top:none;' : '') + '"><span style="color:#000;">コメント</span>　' + esc(item.comment || '特に問題無し') + '</div>';
    h += '<div style="display:flex;gap:2mm;height:48mm;overflow:hidden;">' + photosHtml + '</div>';
    h += '</div></div>';
    return h;
  }

  // 全ページHTML生成
  var pagesHtml = '';
  pages.forEach(function(pageItems, pi) {
    var pageHasDefect = pageItems.some(function(it) { return Number(it.defectQty) > 0; });
    var sampVal = pageHasDefect ? '有' : (sampling || '無');
    var sampBg = pageHasDefect ? 'background-color:#ffff00;' : '';

    var itemsHtml = '';
    for (var c = 0; c < 4; c++) {
      itemsHtml += pageItems[c] ? buildCard(pageItems[c]) : '<div></div>';
    }

    pagesHtml += '<div class="pdf-page">';
    pagesHtml += '<div style="position:absolute;top:10mm;right:12mm;font-size:10px;">' + (pi+1) + '/' + totalPages + '</div>';
    pagesHtml += '<div style="text-align:center;font-size:17px;font-weight:700;letter-spacing:2px;margin-bottom:3mm;color:#000;">【まいばすけっと検質報告書】</div>';
    pagesHtml += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:3mm;">';
    pagesHtml += '<div style="font-size:11px;font-weight:700;">まいばすけっと株式会社　御中</div>';
    pagesHtml += '<div style="font-size:11px;font-weight:700;">' + esc(center) + '農産センター</div>';
    pagesHtml += '</div>';

    // 品目一覧（フル幅）
    pagesHtml += '<div style="border:1px solid #999;padding:2mm 3mm;margin-bottom:3mm;">';
    pagesHtml += '<div style="font-size:9px;font-weight:700;margin-bottom:1.5mm;">今週の検質品目</div>';
    pagesHtml += '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:0 4mm;font-size:8px;line-height:1.6;">' + itemListHtml + '</div>';
    pagesHtml += '</div>';

    // サマリーバー
    pagesHtml += '<div style="display:flex;border:1px solid #999;margin-bottom:4mm;">';
    pagesHtml += '<div style="padding:1.5mm 3mm;border-right:1px solid #999;flex:2;"><div style="font-size:8px;color:#000;">検質日</div><div style="font-weight:700;font-size:13px;">' + date + '</div></div>';
    pagesHtml += '<div style="padding:1.5mm 3mm;border-right:1px solid #999;flex:2;"><div style="font-size:8px;color:#000;">報告者</div><div style="font-weight:700;font-size:13px;">' + esc(staff) + '</div></div>';
    pagesHtml += '<div style="padding:2.5mm 4mm;border-right:1px solid #999;text-align:center;flex:1.2;' + sampBg + '"><div style="font-size:8px;color:#000;">抜取有無</div><div style="font-weight:700;font-size:13px;">' + sampVal + '</div></div>';
    pagesHtml += '<div style="padding:2.5mm 4mm;border-right:1px solid #999;text-align:center;flex:1.5;"><div style="font-size:8px;color:#000;">検質数計</div><div style="font-weight:700;font-size:11px;">' + totalInsp + ' ps</div></div>';
    pagesHtml += '<div style="padding:2.5mm 4mm;border-right:1px solid #999;text-align:center;flex:1.5;"><div style="font-size:8px;color:#000;">対象数計</div><div style="font-weight:700;font-size:11px;">' + totalArrival.toLocaleString() + ' ps</div></div>';
    pagesHtml += '<div style="padding:2.5mm 4mm;text-align:center;flex:1.5;"><div style="font-size:8px;color:#000;">検質率</div><div style="font-weight:700;font-size:11px;">' + inspRate + '%</div></div>';
    pagesHtml += '</div>';

    // 品目カード2×2
    pagesHtml += '<div style="display:grid;grid-template-columns:1fr 1fr;grid-template-rows:1fr 1fr;gap:4mm;flex:1;">' + itemsHtml + '</div>';

    // フッター
    pagesHtml += '<div style="margin-top:3mm;border-top:1px solid #ccc;padding-top:2mm;display:flex;justify-content:space-between;font-size:8px;color:#000;">';
    pagesHtml += '<span>' + esc(center) + '農産センター　TEL: ' + tel + '　FAX: ' + fax + '</span>';
    pagesHtml += '<span>' + (pi+1) + ' / ' + totalPages + '</span>';
    pagesHtml += '</div>';
    pagesHtml += '</div>';
  });

  // 不良レポートページ
  var defectItems = items.filter(function(it) { return Number(it.defectQty) > 0; });
  if (defectItems.length > 0) {
    pagesHtml += '<div class="pdf-page">';
    pagesHtml += '<div style="text-align:center;font-size:17px;font-weight:700;letter-spacing:2px;margin-bottom:3mm;color:#000;">【まいばすけっと検質不良レポート】</div>';
    pagesHtml += '<div style="font-size:11px;font-weight:700;margin-bottom:4mm;">まいばすけっと株式会社　御中</div>';

    pagesHtml += '<div style="display:flex;border:1px solid #999;margin-bottom:5mm;">';
    pagesHtml += '<div style="padding:2mm 4mm;border-right:1px solid #999;flex:1.5;"><div style="font-size:8px;color:#000;">検質日</div><div style="font-weight:700;font-size:13px;">' + date + '</div></div>';
    pagesHtml += '<div style="padding:2mm 4mm;border-right:1px solid #999;flex:1.5;"><div style="font-size:8px;color:#000;">センター</div><div style="font-weight:700;font-size:13px;">' + esc(center) + '農産センター</div></div>';
    pagesHtml += '<div style="padding:2mm 4mm;flex:1.5;"><div style="font-size:8px;color:#000;">報告者</div><div style="font-weight:700;font-size:13px;">' + esc(staff) + '</div></div>';
    pagesHtml += '</div>';

    defectItems.forEach(function(item) {
      var dq = Number(item.defectQty) || 0;
      var iq = Number(item.inspQty) || 0;
      var aq = Number(item.arrivalQty) || 0;
      var rate = iq > 0 ? (Math.round(dq / iq * 1000) / 10) : 0;
      var reason = item.defectReason || '';
      if (reason === 'その他（手入力）' && item.defectReasonText) reason = 'その他: ' + item.defectReasonText;

      pagesHtml += '<div style="border:1.5px solid #c0392b;border-radius:4px;overflow:hidden;margin-bottom:4mm;">';
      pagesHtml += '<div style="background:#c0392b;color:#fff;padding:2mm 3mm;font-size:11px;font-weight:700;">⚠ ' + esc(item.name) + '</div>';
      pagesHtml += '<div style="padding:3mm;display:flex;gap:4mm;">';

      // 左：情報
      pagesHtml += '<div style="flex:1;">';
      pagesHtml += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:2mm;margin-bottom:2mm;font-size:9px;">';
      pagesHtml += '<div><span style="color:#000;font-size:7.5px;">仕入先</span><br><b>' + esc(item.supplier || '-') + '</b></div>';
      pagesHtml += '<div><span style="color:#000;font-size:7.5px;">産地</span><br><b>' + esc(item.origin || '-') + '</b></div>';
      pagesHtml += '<div><span style="color:#000;font-size:7.5px;">入荷数</span><br><b>' + aq + ' ps</b></div>';
      pagesHtml += '<div><span style="color:#000;font-size:7.5px;">検質数</span><br><b>' + iq + ' ps</b></div>';
      pagesHtml += '<div><span style="color:#000;font-size:7.5px;">不良数</span><br><b style="color:#c0392b;">' + dq + ' ps</b></div>';
      pagesHtml += '<div><span style="color:#000;font-size:7.5px;">不良率</span><br><b style="color:#c0392b;">' + rate + '%</b></div>';
      pagesHtml += '</div>';
      if (reason) pagesHtml += '<div style="background:#fff5f5;border-radius:3px;padding:2mm;font-size:9px;margin-bottom:2mm;"><b>不良理由:</b> ' + esc(reason) + '</div>';
      pagesHtml += '<div style="font-size:9px;color:#000;">コメント: ' + esc(item.comment || '') + '</div>';
      pagesHtml += '</div>';

      // 右：不良写真
      pagesHtml += '<div style="display:flex;gap:2mm;min-width:80mm;">';
      var dp = item.defectPhotos || [];
      for (var p = 0; p < 2; p++) {
        if (dp[p]) {
          pagesHtml += '<div style="flex:1;height:48mm;border:1px solid #ccc;border-radius:3px;overflow:hidden;"><img src="' + dp[p] + '" style="width:100%;height:100%;object-fit:contain;display:block;background:#f9f9f9;"></div>';
        } else {
          pagesHtml += '<div style="flex:1;height:48mm;border:1px dashed #ccc;border-radius:3px;display:flex;align-items:center;justify-content:center;font-size:8px;color:#aaa;">写真' + (p+1) + '</div>';
        }
      }
      pagesHtml += '</div>';
      pagesHtml += '</div></div>';
    });

    pagesHtml += '<div style="margin-top:auto;border-top:1px solid #ccc;padding-top:2mm;display:flex;justify-content:space-between;font-size:8px;color:#000;">';
    pagesHtml += '<span>' + esc(center) + '農産センター　TEL: ' + tel + '　FAX: ' + fax + '</span>';
    pagesHtml += '</div>';
    pagesHtml += '</div>';
  }

  // 印刷ウィンドウを開く
  var delivDate = new Date(new Date(date).getTime() + 86400000);
  var dateStr = delivDate.toISOString().slice(0,10).replace(/-/g, '');
  var fileName = '【' + center + '農産】まいばすけっと検質報告書_' + dateStr + '店着';

  var fullHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
    '<title>' + fileName + '</title>' +
    '<style>' +
    '* { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; box-sizing: border-box; margin: 0; padding: 0; }' +
    'body { font-family: "Hiragino Sans","Meiryo",sans-serif; background:#fff; }' +
    '.pdf-page { width:210mm; height:297mm; padding:8mm 10mm 10mm; position:relative; display:flex; flex-direction:column; page-break-after:always; overflow:hidden; box-sizing:border-box; }' +
    '@media print { .pdf-page { page-break-after:always; height:297mm; } @page { size:A4 portrait; margin:0; } }' +
    '</style></head><body>' +
    pagesHtml +
    '<div style="position:fixed;top:10px;right:10px;z-index:9999;display:flex;gap:8px">' +
    '<button onclick="window.print()" style="padding:10px 20px;background:#1a5c2e;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer;font-family:inherit">🖨 印刷</button>' +
    '<button onclick="window.close()" style="padding:10px 16px;background:#fff;color:#333;border:1px solid #ddd;border-radius:8px;font-size:14px;cursor:pointer;font-family:inherit">✕ 閉じる</button>' +
    '</div>' +
    '<style>@media print{div[style*="position:fixed"]{display:none!important}}<\/style>' +
    '</body></html>';

  if (doPrint) {
    // 印刷ウィンドウのみ（Drive保存はしない）
    var win = window.open('', '_blank');
    win.document.write(fullHtml);
    win.document.close();

    return fileName;
  }

  // プレビュー用: HTMLを返す
  return fullHtml;
}

// GASにHTMLを送信してDriveにPDF保存
function savePdfToDrive(center, fileName, html) {
  if (typeof GAS_URL === 'undefined' || !GAS_URL) return;

  var payload = JSON.stringify({
    action: 'savePdfHtml',
    center: center,
    fileName: fileName + '.pdf',
    html: html
  });

  // POST送信（画像含むため大きいデータ）
  fetch(GAS_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: payload,
    redirect: 'follow'
  }).then(function(r) { return r.json(); })
    .then(function(res) {
      if (res.success) {
        if (typeof toast === 'function') toast('Driveに保存しました', 'success');
        console.log('PDF saved to Drive:', res.data.fileName);
      } else {
        console.warn('Drive save error:', res.error);
      }
    }).catch(function(e) {
      console.warn('Drive save failed:', e.message);
    });
}

function esc(s) {
  if (!s) return '';
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
