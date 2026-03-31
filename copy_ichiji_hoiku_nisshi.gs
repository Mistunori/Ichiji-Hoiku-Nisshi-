/**
 * 一時保育日誌 - マスターシートを60枚縦持ちでコピーし、月別シートを作成する
 *
 * 前提:
 *   - マスターシート名: "マスター"
 *   - コピー範囲: 行1〜13（13行）
 *   - G5セル: 月を表す数字（例: 4 → シート名「4月」）
 *   - コピー枚数: 60枚（縦持ち = 13行 × 60 = 780行）
 */

function copyIchijiHoikuNisshi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('マスター');

  if (!masterSheet) {
    SpreadsheetApp.getUi().alert('「マスター」シートが見つかりません。\nシート名を確認してください。');
    return;
  }

  // G5から月（数字のみ）を取得
  const monthValue = masterSheet.getRange('G5').getValue();
  const month = parseInt(monthValue, 10);

  if (isNaN(month) || month < 1 || month > 12) {
    SpreadsheetApp.getUi().alert(
      'G5の値が正しくありません。\n月を表す数字（1〜12）が入力されているか確認してください。\n現在の値: ' + monthValue
    );
    return;
  }

  const sheetName = month + '月';

  // 既存シートの確認
  let targetSheet = ss.getSheetByName(sheetName);
  if (targetSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '確認',
      '「' + sheetName + '」シートは既に存在します。\n上書きしますか？',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) return;

    // 既存シートをクリア
    targetSheet.clearContents();
    targetSheet.clearFormats();
    // 必要な行数を確保
    const needed = 13 * 60;
    if (targetSheet.getMaxRows() < needed) {
      targetSheet.insertRowsAfter(targetSheet.getMaxRows(), needed - targetSheet.getMaxRows());
    }
  } else {
    // 新規シートを作成（行数を最初から確保）
    targetSheet = ss.insertSheet(sheetName);
    const needed = 13 * 60;
    if (targetSheet.getMaxRows() < needed) {
      targetSheet.insertRowsAfter(targetSheet.getMaxRows(), needed - targetSheet.getMaxRows());
    }
  }

  // マスターの範囲（1:13）を取得
  const numCols = masterSheet.getLastColumn();
  const masterRange = masterSheet.getRange(1, 1, 13, numCols);

  // 60枚分、縦持ちで貼り付け
  for (let i = 0; i < 60; i++) {
    const destRow = i * 13 + 1;
    const destRange = targetSheet.getRange(destRow, 1, 13, numCols);
    masterRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  }

  // 列幅をマスターに合わせる
  for (let col = 1; col <= numCols; col++) {
    const width = masterSheet.getColumnWidth(col);
    targetSheet.setColumnWidth(col, width);
  }

  // 行の高さをマスターに合わせる
  for (let row = 1; row <= 13; row++) {
    const height = masterSheet.getRowHeight(row);
    for (let i = 0; i < 60; i++) {
      targetSheet.setRowHeight(i * 13 + row, height);
    }
  }

  SpreadsheetApp.getUi().alert(
    '完了',
    '「' + sheetName + '」シートを作成しました。\n（マスター × 60枚、縦持ち）',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


/**
 * スプレッドシートを開いたときにメニューを追加する
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('一時保育日誌')
    .addItem('月別シートを作成（60枚コピー）', 'copyIchijiHoikuNisshi')
    .addToUi();
}
