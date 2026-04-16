// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 設定値（環境を変えるときはここだけ修正）
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
const SPREADSHEET_ID = "1a07gF0kXMNKNufTzaha0UhttuoF14svy7RgTeXfoaLE";


/**
 * メイン処理：「単価一括登録」→「単価履歴」への転記
 *
 * 処理の流れ:
 *   1. 「単価一括登録」シートから未転記の行を取得
 *   2. 「単価履歴」シートの既存データを読み込み
 *   3. 商社×品名の組み合わせを展開し、重複を除いて履歴に追加
 *   4. 数式を新しい行にコピー（オートフィル）
 *   5. 処理した行に転記済フラグ（N列=1）を立てる
 *
 * 注意:
 *   - E,G,H,J,N,P列はオートフィルの数式が最終値を決める
 *     （スクリプトが書き込む値は数式に上書きされる）
 *   - 安全に処理できる目安は1回あたり入力100行（展開後〜650行）程度
 */
function processNewEntries() {
  const startTime = new Date();

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const bulkRegisterSheet = ss.getSheetByName('単価一括登録');
  const priceHistorySheet = ss.getSheetByName('単価履歴');

  // ──────────────────────────────────────────
  // STEP1: 「単価一括登録」シートから未転記の行を取得
  // ──────────────────────────────────────────
  const bulkAllData = bulkRegisterSheet.getDataRange().getValues();
  const bulkRows = bulkAllData.slice(1); // ヘッダー行を除く

  // N列（14列目, index=13）が 1 でない行 ＝ 未転記
  // != で比較：スプレッドシートが数値1を文字列"1"で返す場合にも対応
  const unprocessedRows = [];
  for (let i = 0; i < bulkRows.length; i++) {
    if (bulkRows[i][13] != 1) {
      unprocessedRows.push({
        data: bulkRows[i],
        sheetRowNumber: i + 2  // ヘッダー=1行目なので、データは2行目から
      });
    }
  }

  if (unprocessedRows.length === 0) {
    Logger.log('新しい登録データはありません。');
    return;
  }

  // ──────────────────────────────────────────
  // STEP2: 「単価履歴」シートの既存データを読み込み
  //        → 重複チェック用リストと前回単価の一覧をメモリ上に作る
  //        ※ ここで1回だけ全件読み込むことで、以降の通信を0にしている
  // ──────────────────────────────────────────
  const historyData = priceHistorySheet.getDataRange().getValues().slice(1);

  // 「日付|処理列1|品名|商社」の組み合わせリスト → 重複チェックに使う
  const processedEntries = buildProcessedEntries(historyData);

  // 「メーカー:商社+品名」→「最新の単価」の対応表 → 前回単価の取得に使う
  const previousPriceMap = buildPreviousPriceMap(historyData);

  // タイムゾーン（日付変換に必要。ループ外で1回だけ取得）
  const timezone = Session.getScriptTimeZone();

  // ──────────────────────────────────────────
  // STEP3: 商社×品名を展開して履歴データを組み立てる
  //        ※ このループ内ではスプレッドシートへの通信は一切行わない
  // ──────────────────────────────────────────
  const newHistoryEntries = [];       // 履歴シートに追加する行データ
  const processedBulkRowNumbers = []; // 転記済フラグを立てる行番号

  unprocessedRows.forEach(({ data: rowData, sheetRowNumber }) => {
    const [
      id, dateObj, manufacturer, traders, processingCol1, items, processingCol2,
      currentPrice, priceChange, previousPrice, spot, spotPeriod, timestamp
    ] = rowData;

    const formattedDate = Utilities.formatDate(dateObj, timezone, 'yyyy/MM/dd');

    // カンマ区切りの商社・品名を個別に分割
    // 例: 「阪和 , 豊通」→ [「阪和」,「豊通」]
    const traderList = traders.split(',').map(tr => tr.trim());
    const itemList = items.split(',').map(it => it.trim());

    let isRowProcessed = false;

    // 全ての商社×品名の組み合わせをループ
    // 例: 2商社×3品名 = 6通り → 履歴に6行追加される
    traderList.forEach(trader => {
      itemList.forEach(itemName => {

        // 重複チェック用のキー（「|」区切りで値の境界を明確にする）
        // ※ 展開後の個別商社で作る（入力行のカンマ区切り全体ではなく）
        const key = formattedDate + '|' + manufacturer + ':' + trader + '|' + itemName + '|' + trader;

        if (!processedEntries.has(key)) {
          // 「メーカー:商社+品名」で前回単価を検索
          // ※ 展開後の個別商社で作る（入力行のカンマ区切り全体ではなく）
          const composedCol2 = manufacturer + ':' + trader + itemName;
          const previousPriceValue = previousPriceMap.get(composedCol2) || 0;
          const newPrice = parseFloat(previousPriceValue) + parseFloat(priceChange || 0);

          // 履歴シートに追加する1行分のデータ
          // ※ E,G,H,J,N,P列はこの後のオートフィルで数式に上書きされる
          newHistoryEntries.push([
            Utilities.getUuid(),                 // A: ユニークID
            formattedDate,                       // B: 日付
            manufacturer,                        // C: メーカー
            trader,                              // D: 商社
            manufacturer + ':' + trader,         // E: 処理列1   ← 数式で上書きされる
            itemName,                            // F: 品名
            composedCol2,                        // G: 処理列2   ← 数式で上書きされる
            newPrice,                            // H: 今回単価   ← 数式で上書きされる
            priceChange,                         // I: 単価変動
            previousPriceValue,                  // J: 前回単価   ← 数式で上書きされる
            spot,                                // K: スポット
            spotPeriod,                          // L: スポット期間
            '',                                  // M: 備考
            '',                                  // N: 最新フラグ ← 数式で上書きされる
            timestamp,                           // O: タイムスタンプ
            manufacturer + ':' + trader + itemName // P: 検索列   ← 数式で上書きされる
          ]);

          // 同じバッチ内の後続行が、この行の単価を「前回単価」として使えるよう更新
          previousPriceMap.set(composedCol2, newPrice);

          // この組み合わせを「処理済み」に追加（同バッチ内の重複防止）
          processedEntries.add(key);
          isRowProcessed = true;
        }
      });
    });

    // 1つでも新規の組み合わせがあれば、この入力行を転記済にする
    if (isRowProcessed) {
      processedBulkRowNumbers.push(sheetRowNumber);
    }
  });

  // ──────────────────────────────────────────
  // STEP4: 履歴シートに一括書き込み → 数式をコピー
  // ──────────────────────────────────────────
  if (newHistoryEntries.length > 0) {
    // 最終行の下に全件まとめて書き込み（通信1回）
    priceHistorySheet
      .getRange(priceHistorySheet.getLastRow() + 1, 1, newHistoryEntries.length, newHistoryEntries[0].length)
      .setValues(newHistoryEntries);

    // 2行目の数式を、今回追加した行だけにコピー
    autoFillFormulas(priceHistorySheet, newHistoryEntries.length);
  }

  // ──────────────────────────────────────────
  // STEP5: 転記済フラグを一括更新
  // ──────────────────────────────────────────
  if (processedBulkRowNumbers.length > 0) {
    batchUpdateTransferFlags(bulkRegisterSheet, processedBulkRowNumbers);
  }

  const elapsed = (new Date() - startTime) / 1000;
  Logger.log(`処理完了: ${newHistoryEntries.length}件追加, ${processedBulkRowNumbers.length}行転記済, ${elapsed}秒`);
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 以下、メイン処理から呼ばれるサブ関数
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━


/**
 * 「単価履歴」の全データから、登録済みの組み合わせリストを作る
 *
 * キーの形式: 「日付|処理列1|品名|商社」
 *   例: 「2026/04/16|中山製鋼:阪和|ＨＳ２|阪和」
 *
 * このリストに含まれるキーは「既に履歴にある」ので、二重登録を防げる
 */
function buildProcessedEntries(historyData) {
  const processed = new Set();
  const tz = Session.getScriptTimeZone();
  historyData.forEach(row => {
    // 日付がDate型の場合は yyyy/MM/dd にフォーマットして統一
    const dateStr = (row[1] instanceof Date)
      ? Utilities.formatDate(row[1], tz, 'yyyy/MM/dd')
      : String(row[1]);
    const key = dateStr + '|' + row[4] + '|' + row[5] + '|' + row[3];
    processed.add(key);
  });
  return processed;
}


/**
 * 「単価履歴」の全データから、品目ごとの最新単価の一覧を作る
 *
 * キー:   処理列2（G列）＝「メーカー:商社+品名」 例: 「中山製鋼:阪和ＨＳ２」
 * 値:     その品目の最新の単価（H列の値）
 *
 * 上から順に読むので、同じキーがあれば後（＝新しい方）で上書きされる
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
 * 転記済フラグ（N列）を一括更新する
 *
 * やっていること:
 *   1. N列を全行まとめて読み込む
 *   2. 対象の行だけメモリ上で「1」に変更
 *   3. 全行まとめて書き戻す
 * → 30行処理しても通信は読み1回＋書き1回の計2回で済む
 */
function batchUpdateTransferFlags(bulkRegisterSheet, rowNumbers) {
  if (rowNumbers.length === 0) return;

  const lastRow = bulkRegisterSheet.getLastRow();
  const flagRange = bulkRegisterSheet.getRange(1, 14, lastRow, 1); // N列 = 14列目
  const flagValues = flagRange.getValues();

  const rowNumberSet = new Set(rowNumbers);
  for (let i = 0; i < flagValues.length; i++) {
    if (rowNumberSet.has(i + 1)) {
      flagValues[i][0] = 1;
    }
  }

  flagRange.setValues(flagValues);
  Logger.log(`転記済フラグを一括更新: ${rowNumbers.length}行`);
}


/**
 * 2行目にある数式を、今回追加した行にコピーする（オートフィル）
 *
 * 対象列: E(処理列1), G(処理列2), H(単価), J(前回単価), N(最新フラグ), P(検索列)
 *
 * ※ 既存の行には触らず、新しく追加した行だけにコピーする
 *   （全行にコピーするとデータ量が多い場合にタイムアウトするため）
 */
function autoFillFormulas(priceHistorySheet, newRowCount) {
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

  // 今回追加した行の開始行を計算（例: 全6000行で5件追加 → 5996行目から）
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
