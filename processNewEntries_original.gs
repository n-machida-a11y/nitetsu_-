function processNewEntries() {

  const spreadSheetId = "1fVClsPMoUzeExsrkIne4_q5QSz4c_v1lGHTN_gqVbSE"; // ★ スプレッドシートのIDを修正
  const bulkRegisterSheetName = '単価一括登録';
  const priceHistorySheetName = '単価履歴';

  const ss = SpreadsheetApp.openById(spreadSheetId);
  const bulkRegisterSheet = ss.getSheetByName(bulkRegisterSheetName);
  const priceHistorySheet = ss.getSheetByName(priceHistorySheetName);

  // 「単価一括登録」シートのデータを取得（ヘッダー行を除き、転記済フラグが0の行のみ）
  const bulkData = bulkRegisterSheet.getDataRange().getValues().slice(1).filter(row => row[13] !== 1);

  if (bulkData.length === 0) {
    Logger.log('新しい登録データはありません。');
    return;
  }

  // 「単価履歴」シートから、すでに登録済みの (日付 + 処理列1 + 品名 + 商社) セットを収集
  const processedEntries = collectProcessedEntries(priceHistorySheet);
  Logger.log('既存の履歴データキー:', processedEntries); // ★ デバッグ用

  const newHistoryEntries = []; // バルクインサート用の配列
  const processedBulkRowIndicesById = new Set(); // IDに基づいて処理済みの行を記録

  // 「単価一括登録」シートの A 列の ID を取得 (ヘッダー除く)
  const allIds = bulkRegisterSheet.getRange('A:A').getValues().flat().slice(1);

  // 未転記の行を順に処理
  bulkData.forEach((rowData, index) => { // index を取得
    const bulkRowNumber = index + 2; // 処理中の「単価一括登録」シートの行番号（デバッグ用）
    Logger.log(`処理中の「単価一括登録」シートの行番号: ${bulkRowNumber}`); // ★ デバッグ用

    const [
      id,
      dateObj, // ★ Dateオブジェクトとして取得
      manufacturer,
      traders,
      processingCol1,
      items,
      processingCol2,
      currentPrice,
      priceChange,
      previousPrice,
      spot,
      spotPeriod,
      timestamp
    ] = rowData;

    // ★ 日付を文字列にフォーマット
    const formattedDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    const traderList = traders.split(',').map(tr => tr.trim());
    const itemList = items.split(',').map(it => it.trim());

    let isRowProcessed = false;

    traderList.forEach(trader => {
      itemList.forEach(itemName => {
        // ★ フォーマットした日付を使用
        const key = formattedDate + processingCol1 + itemName + trader;
        Logger.log(`生成されたキー: ${key}`); // ★ デバッグ用

        if (!processedEntries.has(key)) {
          Logger.log(`新しいキーが見つかりました: ${key}`); // ★ デバッグ用
          // 前回単価の取得
          const composedCol2 = processingCol2 + itemName;
          const previousPriceValue = getPreviousPrice(priceHistorySheet, composedCol2);
          const newPrice = parseFloat(previousPriceValue) + parseFloat(priceChange || 0);

          // 履歴シートに追加する新しい行データ
          const uniqueId = Utilities.getUuid();
          newHistoryEntries.push([
            uniqueId,                // 0:ユニークID
            formattedDate,           // 1:日付 (フォーマット済み)
            manufacturer,            // 2:メーカー
            trader,                  // 3:商社
            manufacturer + ':' + trader, // 4:処理列1に相当
            itemName,                // 5:品名
            composedCol2,            // 6:処理列2 + 品名
            newPrice,                // 7:今回単価
            priceChange,             // 8:単価変動
            previousPriceValue,      // 9:前回単価
            spot,                    // 10:スポット
            spotPeriod,              // 11:スポット期間
            '',                      // 12:備考（空欄）
            '',                      // 13:最新フラグ（空欄）
            timestamp,               // 14:タイムスタンプ
            manufacturer + ':' + trader + itemName // 15:検索列
          ]);

          // 重複登録防止用セットにキーを追加
          processedEntries.add(key);
          isRowProcessed = true;
        } else {
          Logger.log(`既存のキーです: ${key}`); // ★ デバッグ用
        }
      });
    });

    if (isRowProcessed) {
      // 処理済みの「単価一括登録」シートの行の ID を元に、行番号を特定して Set に追加
      const rowIndexToUpdate = allIds.indexOf(id);
      if (rowIndexToUpdate !== -1) {
        processedBulkRowIndicesById.add(rowIndexToUpdate + 2); // +2 はヘッダー行と配列のインデックス調整
        Logger.log(`転記済として記録する行番号 (ID基準): ${rowIndexToUpdate + 2}`);
      } else {
        Logger.log(`ID: ${id} は「単価一括登録」シートで見つかりませんでした。`);
      }
    } else {
      Logger.log(`この行 (ID: ${id}, 行番号: ${bulkRowNumber}) は転記済みと判定されませんでした。`);
    }
  });

  // 履歴シートに新しいデータを一括で書き込み
  if (newHistoryEntries.length > 0) {
    priceHistorySheet.getRange(priceHistorySheet.getLastRow() + 1, 1, newHistoryEntries.length, newHistoryEntries[0].length)
      .setValues(newHistoryEntries);

    // オートフィルを実行
    autoFillFormulas();
  }

  // 「単価一括登録」シートの転記済列をIDに基づいて更新
  processedBulkRowIndicesById.forEach(rowNumberToUpdate => {
    bulkRegisterSheet.getRange(rowNumberToUpdate, 14).setValue(1);
    Logger.log(`転記済フラグを立てた行番号 (ID基準): ${rowNumberToUpdate}`);
  });
}

// ■ 「単価履歴シート」から、(日付+処理列1+品名+商社)の組み合わせを収集
function collectProcessedEntries(priceHistorySheet) {
  const data = priceHistorySheet.getDataRange().getValues().slice(1);
  const processed = new Set();
  data.forEach(row => {
    // 日付、処理列1 (4列目)、品名 (6列目)、商社 (4列目) を結合してキーを作成
    const key = row[1] + row[4] + row[5] + row[3];
    processed.add(key);
  });
  return processed;
}

// ■ 前回単価を取得（priceHistorySheetの末尾から検索）
function getPreviousPrice(priceHistorySheet, composedCol2) {
  const data = priceHistorySheet.getDataRange().getValues().slice(1);
  // 末尾から検索して最初に見つかったcomposedCol2が一致する行の今回単価を返す
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][6] === composedCol2) {
      return data[i][7]; // 今回単価
    }
  }
  return 0; // 見つからない場合は0を返す
}

function autoFillFormulas() {
  const spreadSheetId = "1fVClsPMoUzeExsrkIne4_q5QSz4c_v1lGHTN_gqVbSE"; // ★ スプレッドシートのIDを修正
  const priceHistorySheetName = '単価履歴';
  const ss = SpreadsheetApp.openById(spreadSheetId);
  const priceHistorySheet = ss.getSheetByName(priceHistorySheetName);

  // オートフィルを適用したい列の配列（E, G, H, J, N, P）
  const columnsToAutoFill = ['E', 'G', 'H', 'J', 'N', 'P'];

  columnsToAutoFill.forEach(columnLetter => {
    const firstFormulaCell = columnLetter + '2';
    const formulaRange = priceHistorySheet.getRange(firstFormulaCell);
    const formula = formulaRange.getFormula();

    if (formula) {
      const lastRow = priceHistorySheet.getLastRow();
      if (lastRow > 1) {
        const fillRange = priceHistorySheet.getRange(columnLetter + '2:' + columnLetter + lastRow);
        formulaRange.copyTo(fillRange, { contentsOnly: false }); // 数式と書式をコピー
        Logger.log(`${columnLetter} 列の数式をオートフィルしました。`);
      } else {
        Logger.log(`${columnLetter} 列にはデータがありません。`);
      }
    } else {
      Logger.log(`${columnLetter}2 セルに数式が入力されていません。`);
    }
  });
}
