// ★ スプレッドシートIDを1箇所で管理（変更はここだけ）
const SPREADSHEET_ID = "1fVClsPMoUzeExsrkIne4_q5QSz4c_v1lGHTN_gqVbSE";

/**
 * 単価一括登録 → 単価履歴 転記処理（最適化版）
 *
 * 最適化ポイント:
 *   1. 履歴シートのデータ取得を1回に集約（元: 組み合わせ数×RPC → 1回）
 *   2. メモリ上で processedEntries(Set) と previousPriceMap(Map) を構築
 *   3. 転記済フラグを一括 setValues（元: 行ごとにsetValue）
 *   4. allIds の別途取得を廃止し、行番号を直接追跡
 *   5. autoFillFormulas にシートオブジェクトを引数で渡す
 *   6. Session.getScriptTimeZone() をループ外でキャッシュ
 *   7. キーにセパレータ「|」を追加し、意図しない衝突を防止
 */
function processNewEntries() {
  const startTime = new Date();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const bulkRegisterSheet = ss.getSheetByName('単価一括登録');
  const priceHistorySheet = ss.getSheetByName('単価履歴');

  // ━━━ 一括登録シートのデータを1回で取得 ━━━
  const bulkAllData = bulkRegisterSheet.getDataRange().getValues();
  const bulkHeader = bulkAllData[0];
  const bulkRows = bulkAllData.slice(1);

  // 未転記の行だけ抽出（元の行番号を保持）
  // ★ 修正: != で型の緩い比較（文字列"1"にも対応）
  const unprocessedRows = [];
  for (let i = 0; i < bulkRows.length; i++) {
    if (bulkRows[i][13] != 1) {
      unprocessedRows.push({
        data: bulkRows[i],
        sheetRowNumber: i + 2  // スプレッドシート上の行番号（ヘッダー=1行目なので+2）
      });
    }
  }

  if (unprocessedRows.length === 0) {
    Logger.log('新しい登録データはありません。');
    return;
  }

  // ━━━ 履歴シートのデータを1回だけ取得（最大のボトルネック解消） ━━━
  const historyData = priceHistorySheet.getDataRange().getValues().slice(1);

  // メモリ上で重複チェック用Setを構築
  const processedEntries = buildProcessedEntries(historyData);

  // メモリ上で前回単価Mapを構築（composedCol2 → 最新の今回単価）
  // ★ 末尾が最新なので、順方向で上書きすれば末尾の値が残る
  const previousPriceMap = buildPreviousPriceMap(historyData);

  // ━━━ タイムゾーンを1回だけ取得してキャッシュ ━━━
  const timezone = Session.getScriptTimeZone();

  const newHistoryEntries = [];
  const processedBulkRowNumbers = []; // 転記済にする行番号リスト

  // ━━━ 未転記行を処理（ループ内にシートアクセスなし） ━━━
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
        // ★ セパレータ「|」で意図しないキー衝突を防止
        const key = formattedDate + '|' + processingCol1 + '|' + itemName + '|' + trader;

        if (!processedEntries.has(key)) {
          const composedCol2 = processingCol2 + itemName;

          // ★ メモリ上のMapから前回単価を取得（RPC 0回）
          const previousPriceValue = previousPriceMap.get(composedCol2) || 0;
          const newPrice = parseFloat(previousPriceValue) + parseFloat(priceChange || 0);

          const uniqueId = Utilities.getUuid();
          newHistoryEntries.push([
            uniqueId,                            // 0: ユニークID
            formattedDate,                       // 1: 日付
            manufacturer,                        // 2: メーカー
            trader,                              // 3: 商社
            manufacturer + ':' + trader,         // 4: 処理列1
            itemName,                            // 5: 品名
            composedCol2,                        // 6: 処理列2 + 品名
            newPrice,                            // 7: 今回単価
            priceChange,                         // 8: 単価変動
            previousPriceValue,                  // 9: 前回単価
            spot,                                // 10: スポット
            spotPeriod,                          // 11: スポット期間
            '',                                  // 12: 備考
            '',                                  // 13: 最新フラグ
            timestamp,                           // 14: タイムスタンプ
            manufacturer + ':' + trader + itemName // 15: 検索列
          ]);

          // ★ 同一バッチ内の後続行が参照できるよう、Mapも逐次更新
          previousPriceMap.set(composedCol2, newPrice);

          processedEntries.add(key);
          isRowProcessed = true;
        }
      });
    });

    if (isRowProcessed) {
      processedBulkRowNumbers.push(sheetRowNumber);
    }
  });

  // ━━━ 履歴シートに一括書き込み ━━━
  if (newHistoryEntries.length > 0) {
    priceHistorySheet
      .getRange(priceHistorySheet.getLastRow() + 1, 1, newHistoryEntries.length, newHistoryEntries[0].length)
      .setValues(newHistoryEntries);

    // オートフィル（新規追加分だけにコピー）
    autoFillFormulas(priceHistorySheet, newHistoryEntries.length);
  }

  // ━━━ 転記済フラグを一括更新（元: 行ごとにsetValue → 一括setValues） ━━━
  if (processedBulkRowNumbers.length > 0) {
    batchUpdateTransferFlags(bulkRegisterSheet, processedBulkRowNumbers);
  }

  const elapsed = (new Date() - startTime) / 1000;
  Logger.log(`処理完了: ${newHistoryEntries.length}件追加, ${processedBulkRowNumbers.length}行転記済, ${elapsed}秒`);
}

/**
 * 履歴データから重複チェック用Setを構築（メモリ内処理）
 * ★ セパレータ「|」付きキーに変更
 */
function buildProcessedEntries(historyData) {
  const processed = new Set();
  const tz = Session.getScriptTimeZone(); // ★ ループ外でキャッシュ
  historyData.forEach(row => {
    // row[1]:日付, row[4]:処理列1, row[5]:品名, row[3]:商社
    const dateStr = (row[1] instanceof Date)
      ? Utilities.formatDate(row[1], tz, 'yyyy/MM/dd')
      : String(row[1]);
    const key = dateStr + '|' + row[4] + '|' + row[5] + '|' + row[3];
    processed.add(key);
  });
  return processed;
}

/**
 * 履歴データから前回単価Mapを構築（メモリ内処理）
 * composedCol2（row[6]）をキーに、今回単価（row[7]）を値とする
 * 順方向に走査するので、最後に見つかった値（=最新）が残る
 */
function buildPreviousPriceMap(historyData) {
  const priceMap = new Map();
  historyData.forEach(row => {
    if (row[6] !== '' && row[6] != null) {
      priceMap.set(row[6], row[7]);
    }
  });
  return priceMap;
}

/**
 * 転記済フラグを一括更新
 * 対象行のN列（14列目）に1をセットする
 */
function batchUpdateTransferFlags(bulkRegisterSheet, rowNumbers) {
  if (rowNumbers.length === 0) return;

  // N列全体を一括取得
  const lastRow = bulkRegisterSheet.getLastRow();
  const flagRange = bulkRegisterSheet.getRange(1, 14, lastRow, 1);
  const flagValues = flagRange.getValues();

  // 対象行だけメモリ上で更新
  const rowNumberSet = new Set(rowNumbers);
  for (let i = 0; i < flagValues.length; i++) {
    if (rowNumberSet.has(i + 1)) { // i+1 = スプレッドシートの行番号
      flagValues[i][0] = 1;
    }
  }

  // 一括書き戻し
  flagRange.setValues(flagValues);
  Logger.log(`転記済フラグを一括更新: ${rowNumbers.length}行`);
}

/**
 * 数式のオートフィル（最適化版: 新規追加行だけにコピー）
 * @param {Sheet} priceHistorySheet
 * @param {number} newRowCount - 今回追加した行数（省略時は全行対象）
 */
function autoFillFormulas(priceHistorySheet, newRowCount) {
  // 引数なしで呼ばれた場合の後方互換
  if (!priceHistorySheet) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    priceHistorySheet = ss.getSheetByName('単価履歴');
  }

  const columnsToAutoFill = ['E', 'G', 'H', 'J', 'N', 'P'];
  const lastRow = priceHistorySheet.getLastRow();

  if (lastRow <= 1) {
    Logger.log('オートフィル対象のデータがありません。');
    return;
  }

  // 新規追加行の開始行を算出（指定がなければ2行目から全行）
  const startRow = newRowCount ? lastRow - newRowCount + 1 : 2;

  columnsToAutoFill.forEach(columnLetter => {
    const formulaRange = priceHistorySheet.getRange(columnLetter + '2');
    const formula = formulaRange.getFormula();

    if (formula) {
      const fillRange = priceHistorySheet.getRange(columnLetter + startRow + ':' + columnLetter + lastRow);
      formulaRange.copyTo(fillRange, { contentsOnly: false });
    }
  });
}
