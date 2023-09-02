/*
【実装履歴】
- 20230430
  - 日本株の最新保有株の導出・シート「【株】自動生成：最新保有株」への反映を実装した
    - TODO
      - シミュレート用の処理実装
      - 次回の配当権利確定日の取得と配当権利付最終日の導出，最新保有株への列書き込みの実装
        - 企業別の配当権利確定日を取得できる良い取得元が見つからなかったので実装を保留中
        - https://www.kabutore.biz/haito/haitooti?code=9104
        - https://www.invest-jp.net/yuutai/date_list
      - 米国株用の計算の適用

【Refs】
- https://pineplanter.moo.jp/non-it-salaryman/2022/06/20/gas-header/
  - 列名をインデックスに変換して特定のセルのデータを取得する
- https://into-the-program.com/gas-get-all-data-specified-column-spreadsheet/
  - 列を配列に格納する
- https://auto-worker.com/blog/?p=5785
  - GASでのタイムゾーン設定
- https://qiita.com/nimzo6689/items/99ec2d627ab01a1e3924#javascript-%E3%83%A9%E3%82%A4%E3%83%96%E3%83%A9%E3%83%AA%E3%82%92-gas-%E3%81%A7%E4%BD%BF%E3%81%86%E6%96%B9%E6%B3%95
  - GASでのライブラリ追加
  - 正味期待できない．標準ライブラリで自作するしかない
- https://daily-coding.com/setvalues/
  - データのシートへの書き込み
- https://www.acrovision.jp/service/gas/?p=269
  - 関数を実行するボタンの作成
*/

async function updateStocksHolding() {
  const TODAY = new Date();
  // console.log(TODAY); // 日本時刻になっているか確認
  const REGEX_BLANK = /^[ 　\n]*$/g;
  /*
  スプレッドシートオブジェクトの取得
  */
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  /*
  シート「【株】取引履歴」の情報取得
  */
  const sheetTradeHist = spreadsheet.getSheetByName("【株】取引履歴");
  const lastRowNumTradeHist = sheetTradeHist.getLastRow();
  // ヘッダの取得
  //// カラム名が空白またはスペースのみになっているカラム以降は除外
  let headerTradeHist = sheetTradeHist.getRange("1:1").getValues()[0];
  headerTradeHist = headerTradeHist.slice(0, headerTradeHist.findIndex(v => v.match(REGEX_BLANK)));
  // console.log(headerTradeHist);
  // 処理に必要なカラム名のヘッダにおけるインデックスを取得
  const mapIdxColTradeHist = {
    "コード\n/シンボル": headerTradeHist.findIndex(v => v === "コード\n/シンボル"),
    "取引完了日": headerTradeHist.findIndex(v => v === "取引完了日"),
    "取引採番": headerTradeHist.findIndex(v => v === "取引採番"),
    "【配当取得】\n元URL": headerTradeHist.findIndex(v => v === "【配当取得】\n元URL"),
    "【配当取得】\nCSSSelector": headerTradeHist.findIndex(v => v === "【配当取得】\nCSSSelector"),
    "取引株価": headerTradeHist.findIndex(v => v === "取引株価"),
    "購入後の平均株価\nor売却時の平均株価": headerTradeHist.findIndex(v => v === "購入後の平均株価\nor売却時の平均株価"),
    "取引内容": headerTradeHist.findIndex(v => v === "取引内容"),
    "取引株数": headerTradeHist.findIndex(v => v === "取引株数"),
    "銘柄名": headerTradeHist.findIndex(v => v === "銘柄名"),
    "銘柄カテゴリ": headerTradeHist.findIndex(v => v === "銘柄カテゴリ"),
    "取引状況": headerTradeHist.findIndex(v => v === "取引状況"),
  };
  // console.log(JSON.stringify(mapIdxColTradeHist, null, 2));
  /*
  シート「【株】自動生成：最新保有株」の情報取得
  */
  const sheetStocksHolding = spreadsheet.getSheetByName("【株】自動生成：最新保有株");
  const lastRowNumStocksHolding = sheetStocksHolding.getLastRow();
  // ヘッダの取得
  //// カラム名が空白またはスペースのみになっているカラム以降は除外
  let headerStocksHolding = sheetStocksHolding.getRange("1:1").getValues()[0];
  headerStocksHolding = headerStocksHolding.slice(0, headerStocksHolding.findIndex(v => v.match(REGEX_BLANK)));
  // console.log(headerStocksHolding);
  // 処理に必要なカラム名のヘッダにおけるインデックスを取得
  const mapIdxColStocksHolding = {
    "コード\n/シンボル": headerStocksHolding.findIndex(v => v === "コード\n/シンボル"),
    "銘柄名": headerStocksHolding.findIndex(v => v === "銘柄名"),
    "銘柄カテゴリ": headerStocksHolding.findIndex(v => v === "銘柄カテゴリ"),
    "【配当取得】\n元URL": headerStocksHolding.findIndex(v => v === "【配当取得】\n元URL"),
    "【配当取得】\nCSSSelector": headerStocksHolding.findIndex(v => v === "【配当取得】\nCSSSelector"),
    "現在保有している株の保有開始日": headerStocksHolding.findIndex(v => v === "現在保有している株の保有開始日"),
    "保有株数": headerStocksHolding.findIndex(v => v === "保有株数"),
    "平均株価": headerStocksHolding.findIndex(v => v === "平均株価"),
    "直近の取引日": headerStocksHolding.findIndex(v => v === "直近の取引日"),
    "直近の取引株価": headerStocksHolding.findIndex(v => v === "直近の取引株価"),
    // "次回権利付最終日": headerStocksHolding.findIndex(v => v === "次回権利付最終日"),
    // "次回権利確定日": headerStocksHolding.findIndex(v => v === "次回権利確定日"),
  };

  /*
  シート「【株】取引履歴」のデータをもとに最新の保有株を取得し，シート「【株】自動生成：最新保有株」に上書き書き込み．
  */
  // シート「【株】取引履歴」のデータの読み込み
  //// ヘッダを除く2行目以降を取得
  //// 「取引状況＝完了」の行のみ選択
  const rowsTradeHist = sheetTradeHist
    .getRange(2, 1, lastRowNumTradeHist, headerTradeHist.length)
    .getValues()
    .filter(r => r[mapIdxColTradeHist["取引状況"]] === "完了");
  // console.log(rowsTradeHist.slice(0, 10));
  // console.log(rowsTradeHist.length);
  // コード/シンボルの一意な一覧取得
  const simbolsTradeHist = [
    ...new Set(rowsTradeHist.map(r => r[mapIdxColTradeHist["コード\n/シンボル"]]))
  ].filter(v => !String(v).match(REGEX_BLANK));
  // console.log(simbolsTradeHist);
  let rowsStockHolding = [];
  return await Promise.all(
    simbolsTradeHist.map(symbol =>
      insertRowStockHolding(rowsStockHolding, rowsTradeHist, mapIdxColTradeHist, mapIdxColStocksHolding, symbol)
    )
  )
    .then(_ => {
      // console.log(rowsStockHolding); // この変数にシート「【株】自動生成：最新保有株」に表示させる，カラム順も整ったデータができているはず．
      /*
      シート「【株】自動生成：最新保有株」のデータ入れ直し
      */
      // ヘッダを除く2行目以降の行をクリア
      sheetStocksHolding.getRange(2, 1, lastRowNumStocksHolding, headerStocksHolding.length).clear();
      if (rowsStockHolding.length > 0) {
        // 2行目以降にデータを書き込み
        // console.log(2, 1, rowsStockHolding.length, Object.keys(mapIdxColStocksHolding).length);
        sheetStocksHolding.getRange(2, 1, rowsStockHolding.length, Object.keys(mapIdxColStocksHolding).length).setValues(rowsStockHolding);
      }
    })
    .catch(err => {
      throw err;
    });
}

async function insertRowStockHolding(ret, rowsTradeHist, mapIdxColTradeHist, mapIdxColStocksHolding, symbol) {
  /*
  シンボルごとに最新保有状況を導出して，最新保有株リストに登録する．
  ただし最新の保有株数が0の場合は登録しない．

  - Args
    - ret:Array: 処理結果のinsert先の配列
    - rowsTradeHist:Array: 取引履歴データ
    - mapIdxColTradeHist:Array: シート「【株】取引履歴」のヘッダのカラム名とインデックス対応表
    - mapIdxColTradeHist:Array: シート「【株】自動生成：最新保有株」のヘッダのカラム名とインデックス対応表
    - symbol:string|number: コード/シンボル
  */
  // あるシンボルの取引履歴を抽出し，取引完了日・取引採番の順番にソートする
  const rows = rowsTradeHist
    .filter(r => symbol === r[mapIdxColTradeHist["コード\n/シンボル"]])
    .sort(sortByTrandeDateAndTradeNumAsc(mapIdxColTradeHist["取引完了日"], mapIdxColTradeHist["取引採番"]));
  // console.info(rows);
  // 直近の取引完了日の行を一旦抽出しておく
  const latestRow = rows.reduce((accumulator, row) => (accumulator[mapIdxColTradeHist["取引完了日"]] > row["取引完了日"]) ? accumulator : row);
  let startDateLatestStockHolding = latestRow[mapIdxColTradeHist["取引完了日"]];
  let stockCntHolding = 0;
  for (const row of rows) {
    if (stockCntHolding === 0) {
      // 取引株数の増加量を加味する前が0個だったら保有開始日であるとみなす
      startDateLatestStockHolding = row[mapIdxColTradeHist["取引完了日"]];
    }
    const incAmount = (row[mapIdxColTradeHist["取引内容"]] === "購入") ? row[mapIdxColTradeHist["取引株数"]] : -row[mapIdxColTradeHist["取引株数"]];
    stockCntHolding = stockCntHolding + incAmount;
  }
  // console.info(`${symbol}: ${stockCntHolding}`);
  if (stockCntHolding !== 0) {
    // 先に辞書形式でカラム名とデータを紐づけインデックスでカラム順整えた後で，insertする．
    const stockHoldingData = {
      "コード\n/シンボル": symbol,
      "銘柄名": latestRow[mapIdxColTradeHist["銘柄名"]], // 直近の取引完了日の行から取得
      "銘柄カテゴリ": latestRow[mapIdxColTradeHist["銘柄カテゴリ"]], // 直近の取引完了日の行から取得
      "【配当取得】\n元URL": latestRow[mapIdxColTradeHist["【配当取得】\n元URL"]], // 直近の取引完了日の行から取得
      "【配当取得】\nCSSSelector": latestRow[mapIdxColTradeHist["【配当取得】\nCSSSelector"]], // 直近の取引完了日の行から取得
      "現在保有している株の保有開始日": startDateLatestStockHolding,
      "保有株数": stockCntHolding,
      "平均株価": latestRow[mapIdxColTradeHist["購入後の平均株価\nor売却時の平均株価"]], // 直近の取引完了日の行から取得
      "直近の取引日": latestRow[mapIdxColTradeHist["取引完了日"]], // 直近の取引完了日の行から取得
      "直近の取引株価": latestRow[mapIdxColTradeHist["取引株価"]], // 直近の取引完了日の行から取得
      // "次回権利付最終日": ,
      // "次回権利確定日": ,
    };
    // console.info(stockHoldingData);

    let rowStockHolding = [];
    try {
      for (let i = 0; i < Object.keys(mapIdxColStocksHolding).length; i++) {
        const colNames = Object.keys(mapIdxColStocksHolding).filter(k => mapIdxColStocksHolding[k] === i);
        // console.info(colNames);
        if (colNames.length === 0) {
          throw Error("`mapIdxColStocksHolding`の宣言時における，シート「【株】自動生成：最新保有株」のカラム名の取得が十分にできていない可能性があります．");
        } else if (colNames.length > 1) {
          throw Error("`mapIdxColStocksHolding`の宣言時における，シート「【株】自動生成：最新保有株」のカラム名の取得が適切でない可能性があります．")
        }
        rowStockHolding.push(stockHoldingData[colNames[0]]);
      }
    } catch (err) {
      throw err;
    }
    // console.info(rowStockHolding);
    // return rowStockHolding;
    ret.push(rowStockHolding);
  } else {
    console.info(`{"シンボル": "${symbol}", "銘柄名": "${latestRow[mapIdxColTradeHist["銘柄名"]]}"}の最新保有株数は0でした．`);
  }
}

function sortByTrandeDateAndTradeNumAsc(idxColTradeDate, idxColTradeNum) {
  /*
  「ORDER BY 取引完了日, 取引採番 ASC」のソートに用いる．
  - Args
    - idxColTradeDate:number: ソートする比較対象のデータにおける取引完了日のカラムインデックス
    - idxColTradeNum:number: ソートする比較対象のデータにおける取引採番のカラムインデックス
  - Refs
    - https://dev.to/markbdsouza/js-sort-an-array-of-objects-on-multiple-columns-keys-2bj1
    - https://www.benmvp.com/blog/quick-way-sort-javascript-array-multiple-fields/
    - https://www.tutorialspoint.com/accessing-nested-javascript-objects-with-string-key
    - https://stackoverflow.com/questions/33016087/js-sort-empty-to-end
  */
  function validate(a, b) {
    /*
    - Args
      - a:Array:比較対象のデータ
      - b:Array:比較対象のデータ
    */
    if (!a[idxColTradeDate] || !a[idxColTradeNum] || !b[idxColTradeDate] || !b[idxColTradeNum]) {
      throw Error("「ORDER BY 取引完了日, 取引採番 ASC」のソートするために，取引完了日と取引採番が埋められている必要があります．");
    }
    const aTradeDate = new Date(a[idxColTradeDate]);
    const bTradeDate = new Date(b[idxColTradeDate]);
    return (aTradeDate - bTradeDate) || (a[idxColTradeNum] - b[idxColTradeNum]);
  }
  return validate;
}
