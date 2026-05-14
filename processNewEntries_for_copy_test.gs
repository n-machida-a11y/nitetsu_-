// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// コピースプシでの動作検証用
//   - processNewEntries にタイミング計測を仕込んだ版
//   - setupDemoData() でテストデータを一括投入
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const SPREADSHEET_ID = "1a07gF0kXMNKNufTzaha0UhttuoF14svy7RgTeXfoaLE";


function processNewEntries() {
  const startTime = new Date();
  const t = (label) => Logger.log(`[+${((new Date() - startTime) / 1000).toFixed(2)}s] ${label}`);

  t('開始');
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const bulkRegisterSheet = ss.getSheetByName('単価一括登録');
  const priceHistorySheet = ss.getSheetByName('単価履歴');
  t('シート取得');

  const bulkAllData = bulkRegisterSheet.getDataRange().getValues();
  const bulkRows = bulkAllData.slice(1);
  t(`単価一括登録読込 (${bulkRows.length}行)`);

  const unprocessedRows = [];
  for (let i = 0; i < bulkRows.length; i++) {
    if (bulkRows[i][13] != 1) {
      unprocessedRows.push({ data: bulkRows[i], sheetRowNumber: i + 2 });
    }
  }
  t(`未転記抽出 (${unprocessedRows.length}行)`);

  if (unprocessedRows.length === 0) {
    Logger.log('新しい登録データはありません。');
    return;
  }

  const historyData = priceHistorySheet.getDataRange().getValues().slice(1);
  t(`単価履歴読込 (${historyData.length}行)`);

  const processedEntries = buildProcessedEntries(historyData);
  const previousPriceMap = buildPreviousPriceMap(historyData);
  t('Set/Map構築');

  const timezone = Session.getScriptTimeZone();
  const newHistoryEntries = [];
  const processedBulkRowNumbers = [];

  unprocessedRows.forEach(({ data: rowData, sheetRowNumber }) => {
    const [
      id, dateObj, manufacturer, traders, processingCol1, items, processingCol2,
      currentPrice, priceChange, previousPrice, spot, spotPeriod, timestamp
    ] = rowData;

    const formattedDate = Utilities.formatDate(dateObj, timezone, 'yyyy/MM/dd');
    const traderList = traders.split(',').map(tr => tr.trim());
    const itemList = items.split(',').map(it => it.trim());
    let isRowProcessed = false;

    traderList.forEach(trader => {
      itemList.forEach(itemName => {
        const key = formattedDate + '|' + manufacturer + ':' + trader + '|' + itemName + '|' + trader;
        if (!processedEntries.has(key)) {
          const composedCol2 = manufacturer + ':' + trader + itemName;
          const previousPriceValue = previousPriceMap.get(composedCol2) || 0;
          const newPrice = parseFloat(previousPriceValue) + parseFloat(priceChange || 0);
          newHistoryEntries.push([
            Utilities.getUuid(), formattedDate, manufacturer, trader,
            manufacturer + ':' + trader, itemName, composedCol2, newPrice,
            priceChange, previousPriceValue, spot, spotPeriod, '', '',
            timestamp, manufacturer + ':' + trader + itemName
          ]);
          previousPriceMap.set(composedCol2, newPrice);
          processedEntries.add(key);
          isRowProcessed = true;
        }
      });
    });
    if (isRowProcessed) processedBulkRowNumbers.push(sheetRowNumber);
  });
  t(`組み立て (${newHistoryEntries.length}件追加予定)`);

  if (newHistoryEntries.length > 0) {
    priceHistorySheet
      .getRange(priceHistorySheet.getLastRow() + 1, 1, newHistoryEntries.length, newHistoryEntries[0].length)
      .setValues(newHistoryEntries);
    t('履歴シート setValues');

    autoFillFormulas(priceHistorySheet, newHistoryEntries.length);
    t('autoFillFormulas 完了');
  }

  if (processedBulkRowNumbers.length > 0) {
    batchUpdateTransferFlags(bulkRegisterSheet, processedBulkRowNumbers);
    t('転記済フラグ更新');
  }

  const elapsed = (new Date() - startTime) / 1000;
  Logger.log(`処理完了: ${newHistoryEntries.length}件追加, ${processedBulkRowNumbers.length}行転記済, ${elapsed}秒`);
}


function buildProcessedEntries(historyData) {
  const processed = new Set();
  const tz = Session.getScriptTimeZone();
  historyData.forEach(row => {
    const dateStr = (row[1] instanceof Date)
      ? Utilities.formatDate(row[1], tz, 'yyyy/MM/dd')
      : String(row[1]);
    processed.add(dateStr + '|' + row[4] + '|' + row[5] + '|' + row[3]);
  });
  return processed;
}


function buildPreviousPriceMap(historyData) {
  const priceMap = new Map();
  historyData.forEach(row => {
    if (row[6] !== '' && row[6] != null) {
      priceMap.set(row[6], row[7]);
    }
  });
  return priceMap;
}


function batchUpdateTransferFlags(bulkRegisterSheet, rowNumbers) {
  if (rowNumbers.length === 0) return;
  const lastRow = bulkRegisterSheet.getLastRow();
  const flagRange = bulkRegisterSheet.getRange(1, 14, lastRow, 1);
  const flagValues = flagRange.getValues();
  const rowNumberSet = new Set(rowNumbers);
  for (let i = 0; i < flagValues.length; i++) {
    if (rowNumberSet.has(i + 1)) flagValues[i][0] = 1;
  }
  flagRange.setValues(flagValues);
}


function autoFillFormulas(priceHistorySheet, newRowCount) {
  const start = new Date();
  if (!priceHistorySheet) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    priceHistorySheet = ss.getSheetByName('単価履歴');
  }

  const columnsToAutoFill = ['E', 'G', 'H', 'J', 'N', 'P'];
  const lastRow = priceHistorySheet.getLastRow();
  if (lastRow <= 1) return;
  const startRow = newRowCount ? lastRow - newRowCount + 1 : 2;

  columnsToAutoFill.forEach(columnLetter => {
    const colStart = new Date();
    const formulaRange = priceHistorySheet.getRange(columnLetter + '2');
    const formula = formulaRange.getFormula();
    if (formula) {
      const fillRange = priceHistorySheet.getRange(columnLetter + startRow + ':' + columnLetter + lastRow);
      formulaRange.copyTo(fillRange, { contentsOnly: false });
    }
    Logger.log(`  ${columnLetter}列 copyTo: ${((new Date() - colStart) / 1000).toFixed(2)}秒`);
  });
  Logger.log(`autoFillFormulas 内訳合計: ${((new Date() - start) / 1000).toFixed(2)}秒`);
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 1回だけ実行：コピースプシにデモデータを投入
//   - 「単価履歴」「単価一括登録」シートが既に存在している前提
//   - 既存データは clear() で消えるので注意
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function setupDemoData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const historySheet = ss.getSheetByName('単価履歴');
  const bulkSheet = ss.getSheetByName('単価一括登録');

  if (!historySheet || !bulkSheet) {
    Logger.log('「単価履歴」または「単価一括登録」シートが見つかりません。先にシートを用意してください。');
    return;
  }

  const historyHeader = ['id','日付','メーカー','商社','処理列1','品名','処理列2','単価','単価上げ下げ','前回単価','スポット','スポット期間','備考','最新フラグ','タイムスタンプ','検索列'];
  const bulkHeader = ['id','日付','メーカー','商社','処理列1','品名','処理列2','単価','単価上げ下げ','前回単価','スポット','スポット期間','タイムスタンプ','転記済'];

  historySheet.clear();
  historySheet.getRange(1, 1, 1, historyHeader.length).setValues([historyHeader]);
  const historyRows = [
    ['hist001', '2026/04/01', '中山製鋼', '阪和', '中山製鋼:阪和', 'ＨＳ２', '中山製鋼:阪和ＨＳ２', 50.0, 0, 50.0, '', '', '', '', '2026/04/01 10:00:00', '中山製鋼:阪和ＨＳ２'],
    ['hist002', '2026/04/01', '中山製鋼', '豊通', '中山製鋼:豊通', 'ＨＳ２', '中山製鋼:豊通ＨＳ２', 48.0, 0, 48.0, '', '', '', '', '2026/04/01 10:00:00', '中山製鋼:豊通ＨＳ２'],
    ['hist003', '2026/04/01', '中山製鋼', '阪和', '中山製鋼:阪和', 'ＨＳ１(単体H1/電特A/HS))', '中山製鋼:阪和ＨＳ１(単体H1/電特A/HS))', 55.0, 0, 55.0, '', '', '', '', '2026/04/01 10:00:00', '中山製鋼:阪和ＨＳ１(単体H1/電特A/HS))'],
    ['hist004', '2026/04/01', '中山製鋼', '豊通', '中山製鋼:豊通', 'ＨＳ１(単体H1/電特A/HS))', '中山製鋼:豊通ＨＳ１(単体H1/電特A/HS))', 53.0, 0, 53.0, '', '', '', '', '2026/04/01 10:00:00', '中山製鋼:豊通ＨＳ１(単体H1/電特A/HS))']
  ];
  historySheet.getRange(2, 1, historyRows.length, historyHeader.length).setValues(historyRows);

  bulkSheet.clear();
  bulkSheet.getRange(1, 1, 1, bulkHeader.length).setValues([bulkHeader]);
  const bulkRows = [
    ['test001', '2026/04/16', '中山製鋼', '阪和', '中山製鋼:阪和', 'ＨＳ２', '中山製鋼:阪和ＨＳ２', '', 0.5, '', '', '', '2026/04/16 10:00:00', ''],
    ['test002', '2026/04/16', '中山製鋼', '阪和 , 豊通', '中山製鋼:阪和 , 豊通', 'ＨＳ１(単体H1/電特A/HS)) , ＨＳ２', '中山製鋼:阪和 , 豊通ＨＳ１(単体H1/電特A/HS)) , ＨＳ２', '', 1.0, '', '', '', '2026/04/16 10:01:00', ''],
    ['test003', '2026/04/16', '中山製鋼', '阪和', '中山製鋼:阪和', 'ＨＳ２', '中山製鋼:阪和ＨＳ２', '', 0.5, '', '', '', '2026/04/16 10:02:00', '']
  ];
  bulkSheet.getRange(2, 1, bulkRows.length, bulkHeader.length).setValues(bulkRows);

  Logger.log(`デモデータ投入完了: 単価履歴=${historyRows.length}行, 単価一括登録=${bulkRows.length}行`);
}
