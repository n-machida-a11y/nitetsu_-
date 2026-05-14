// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// コピースプシでの動作検証用
//   - processNewEntries にタイミング計測を仕込んだ版
//   - setupDemoData() でテストデータを一括投入
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 新コピー（本番からまるごとコピーした検証用スプシ）
const SPREADSHEET_ID = "1cptWC-9wY2s9ClVM94axWVMnHi60AbZcNuVOx1Jp5-8";


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

    updateLatestFlags(priceHistorySheet, historyData, newHistoryEntries);
    t('N列(最新フラグ)更新');
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


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// N列(最新フラグ●)を計算して書き戻す
//   - 元のper-row MAXIFS数式と同じセマンティクス:
//     「(E,F)組合せのうち、自分のOがその組合せの最大Oと一致 → ●」
//   - 数式を使わずGASでO(N)で計算→1回のsetValuesで書き込み
//   - 単価履歴シートにN列の数式が無い前提（事前に値貼り付けで撤去済）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function updateLatestFlags(historySheet, historyData, newHistoryEntries) {
  const allRows = historyData.concat(newHistoryEntries);

  // Dateオブジェクトは === が参照比較になるため、比較用に数値化する
  const toComparable = (v) => {
    if (v === '' || v === null || v === undefined) return null;
    if (v instanceof Date) return v.getTime();
    return v;
  };

  // (E|F) -> 最大Oの「比較可能な値」
  const maxOMap = new Map();
  allRows.forEach(row => {
    const key = (row[4] || '') + '|' + (row[5] || '');
    const oVal = toComparable(row[14]);
    if (oVal === null) return;
    const curr = maxOMap.get(key);
    if (curr === undefined || oVal > curr) {
      maxOMap.set(key, oVal);
    }
  });

  // 各行のN列の値（●か空）
  // 値比較なので、同じ最大Oを持つ行は全部に ● がつく（MAXIFS数式と同セマンティクス）
  const desiredN = allRows.map(row => {
    const key = (row[4] || '') + '|' + (row[5] || '');
    const oVal = toComparable(row[14]);
    const max = maxOMap.get(key);
    return [(oVal !== null && oVal === max) ? '●' : ''];
  });

  // N列(14列目)に一括書き込み
  historySheet.getRange(2, 14, desiredN.length, 1).setValues(desiredN);
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
// 本番スプシ vs 新コピースプシ の N列(最新フラグ●)を id基準で突き合わせ
//   - A列(id)をキーに、両シートの N列を比較
//   - 一致/不一致/片方にだけ存在する行 をカウント
//   - 不一致サンプルを最大20件ログ出力
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function compareLatestFlags() {
  const productionId = '1fVClsPMoUzeExsrkIne4_q5QSz4c_v1lGHTN_gqVbSE';
  const verifyId = '1cptWC-9wY2s9ClVM94axWVMnHi60AbZcNuVOx1Jp5-8';

  const prodSheet = SpreadsheetApp.openById(productionId).getSheetByName('単価履歴');
  const verifySheet = SpreadsheetApp.openById(verifyId).getSheetByName('単価履歴');

  const prodLastRow = prodSheet.getLastRow();
  const verifyLastRow = verifySheet.getLastRow();
  Logger.log(`本番: ${prodLastRow - 1}行, 検証: ${verifyLastRow - 1}行`);

  if (prodLastRow <= 1 || verifyLastRow <= 1) {
    Logger.log('片方のシートが空です');
    return;
  }

  // A列(id)とN列(最新フラグ)だけ取得
  const prodIds = prodSheet.getRange(2, 1, prodLastRow - 1, 1).getValues().map(r => r[0]);
  const prodN = prodSheet.getRange(2, 14, prodLastRow - 1, 1).getValues().map(r => r[0]);
  const verifyIds = verifySheet.getRange(2, 1, verifyLastRow - 1, 1).getValues().map(r => r[0]);
  const verifyN = verifySheet.getRange(2, 14, verifyLastRow - 1, 1).getValues().map(r => r[0]);

  // 検証側を id -> N のMapに
  const verifyMap = new Map();
  for (let i = 0; i < verifyIds.length; i++) {
    verifyMap.set(verifyIds[i], verifyN[i]);
  }

  // 本番をループして比較
  let matchCount = 0;
  let mismatchCount = 0;
  let onlyInProduction = 0;
  const mismatchSamples = [];

  for (let i = 0; i < prodIds.length; i++) {
    const id = prodIds[i];
    if (!verifyMap.has(id)) {
      onlyInProduction++;
      continue;
    }
    const vN = verifyMap.get(id);
    if ((prodN[i] || '') === (vN || '')) {
      matchCount++;
    } else {
      mismatchCount++;
      if (mismatchSamples.length < 20) {
        mismatchSamples.push({
          row: i + 2,
          id: String(id).substring(0, 12),
          prod: prodN[i],
          verify: vN
        });
      }
    }
  }

  // 検証側にしかないidの数
  const prodIdSet = new Set(prodIds);
  let onlyInVerify = 0;
  for (let i = 0; i < verifyIds.length; i++) {
    if (!prodIdSet.has(verifyIds[i])) onlyInVerify++;
  }

  Logger.log(`一致: ${matchCount} / 不一致: ${mismatchCount} / 本番のみ: ${onlyInProduction} / 検証のみ: ${onlyInVerify}`);

  if (mismatchSamples.length > 0) {
    Logger.log(`--- 不一致サンプル (${Math.min(mismatchCount, 20)}件まで表示) ---`);
    mismatchSamples.forEach(m => {
      Logger.log(`  検証行${m.row} id=${m.id}... 本番N='${m.prod}' / 検証N='${m.verify}'`);
    });
  } else if (mismatchCount === 0) {
    Logger.log('★ 両シートで突き合わせた全行のN列が完全一致');
  }
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 本番スプシ vs コピースプシ の読み取り速度ベンチマーク
//   - 書き込みはせず、開く / シート取得 / getValues の各時間を比較するだけ
//   - 本番が極端に遅ければ「本番スプシ自体が重い」ことが確定する
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function benchmarkRead() {
  const targets = [
    { label: '旧コピー(816行)', id: '1a07gF0kXMNKNufTzaha0UhttuoF14svy7RgTeXfoaLE' },
    { label: '新コピー(本番相当)', id: '1cptWC-9wY2s9ClVM94axWVMnHi60AbZcNuVOx1Jp5-8' },
    { label: '本番',              id: '1fVClsPMoUzeExsrkIne4_q5QSz4c_v1lGHTN_gqVbSE' }
  ];

  targets.forEach(({ label, id }) => {
    const t0 = new Date();
    const ss = SpreadsheetApp.openById(id);
    const t1 = new Date();

    const historySheet = ss.getSheetByName('単価履歴');
    const bulkSheet = ss.getSheetByName('単価一括登録');
    const t2 = new Date();

    const historyData = historySheet.getDataRange().getValues();
    const t3 = new Date();

    const bulkData = bulkSheet.getDataRange().getValues();
    const t4 = new Date();

    Logger.log(
      `[${label}] openById=${((t1 - t0) / 1000).toFixed(2)}s, ` +
      `シート取得=${((t2 - t1) / 1000).toFixed(2)}s, ` +
      `単価履歴 getValues=${((t3 - t2) / 1000).toFixed(2)}s (${historyData.length}行), ` +
      `単価一括登録 getValues=${((t4 - t3) / 1000).toFixed(2)}s (${bulkData.length}行), ` +
      `合計=${((t4 - t0) / 1000).toFixed(2)}s`
    );
  });
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// コピースプシにデモデータを追記投入（既存データは触らない）
//   - 「単価履歴」と「単価一括登録」の末尾に行を追加するだけ
//   - ヘッダは既に存在している前提（書き換えない）
//   - 再実行すると同じ内容が重複追加される点に注意
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function setupDemoData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const historySheet = ss.getSheetByName('単価履歴');
  const bulkSheet = ss.getSheetByName('単価一括登録');

  if (!historySheet || !bulkSheet) {
    Logger.log('「単価履歴」または「単価一括登録」シートが見つかりません。先にシートを用意してください。');
    return;
  }

  const historyRows = [
    ['hist001', '2026/04/01', '中山製鋼', '阪和', '中山製鋼:阪和', 'ＨＳ２', '中山製鋼:阪和ＨＳ２', 50.0, 0, 50.0, '', '', '', '', '2026/04/01 10:00:00', '中山製鋼:阪和ＨＳ２'],
    ['hist002', '2026/04/01', '中山製鋼', '豊通', '中山製鋼:豊通', 'ＨＳ２', '中山製鋼:豊通ＨＳ２', 48.0, 0, 48.0, '', '', '', '', '2026/04/01 10:00:00', '中山製鋼:豊通ＨＳ２'],
    ['hist003', '2026/04/01', '中山製鋼', '阪和', '中山製鋼:阪和', 'ＨＳ１(単体H1/電特A/HS))', '中山製鋼:阪和ＨＳ１(単体H1/電特A/HS))', 55.0, 0, 55.0, '', '', '', '', '2026/04/01 10:00:00', '中山製鋼:阪和ＨＳ１(単体H1/電特A/HS))'],
    ['hist004', '2026/04/01', '中山製鋼', '豊通', '中山製鋼:豊通', 'ＨＳ１(単体H1/電特A/HS))', '中山製鋼:豊通ＨＳ１(単体H1/電特A/HS))', 53.0, 0, 53.0, '', '', '', '', '2026/04/01 10:00:00', '中山製鋼:豊通ＨＳ１(単体H1/電特A/HS))']
  ];
  const historyStart = historySheet.getLastRow() + 1;
  historySheet.getRange(historyStart, 1, historyRows.length, historyRows[0].length).setValues(historyRows);

  const bulkRows = [
    ['test001', '2026/04/16', '中山製鋼', '阪和', '中山製鋼:阪和', 'ＨＳ２', '中山製鋼:阪和ＨＳ２', '', 0.5, '', '', '', '2026/04/16 10:00:00', ''],
    ['test002', '2026/04/16', '中山製鋼', '阪和 , 豊通', '中山製鋼:阪和 , 豊通', 'ＨＳ１(単体H1/電特A/HS)) , ＨＳ２', '中山製鋼:阪和 , 豊通ＨＳ１(単体H1/電特A/HS)) , ＨＳ２', '', 1.0, '', '', '', '2026/04/16 10:01:00', ''],
    ['test003', '2026/04/16', '中山製鋼', '阪和', '中山製鋼:阪和', 'ＨＳ２', '中山製鋼:阪和ＨＳ２', '', 0.5, '', '', '', '2026/04/16 10:02:00', '']
  ];
  const bulkStart = bulkSheet.getLastRow() + 1;
  bulkSheet.getRange(bulkStart, 1, bulkRows.length, bulkRows[0].length).setValues(bulkRows);

  Logger.log(`デモデータ追加完了: 単価履歴=${historyRows.length}行を行${historyStart}以降に追加, 単価一括登録=${bulkRows.length}行を行${bulkStart}以降に追加`);
}
