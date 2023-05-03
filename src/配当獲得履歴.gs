/*
【実装履歴】
- 20230430
  - 日本株の最新保有株の算出・シート「【株】自動生成：配当獲得履歴」への反映を実装した
    - TODO
      - シミュレート用の処理実装
      - 米国株用の計算の適用

【Refs】
- https://qiita.com/standard-software/items/16214dc4e64d28952c2d
- https://self-methods.com/gas-urlfetchapp/
  - GASでのWebページのHTMLダウンロード
- https://qiita.com/kairi003/items/06fbf2dc8fb5415c7f60
  - GASでquerySelectorAllでスクレイピング
  - https://github.com/kairi003/gas-html-parser
- https://qiita.com/RyBB/items/c87af2413c34f9367d00
  - textContext/innerHTML/innerTextの違い
*/

async function appendDividendHist() {
  const TODAY = new Date();
  // console.log(TODAY); // 日本時刻になっているか確認
  const REGEX_BLANK = /^[ 　\n]*$/g;
  /*
  スプレッドシートオブジェクトの取得
  */
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
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
  };
  /*
  シート「【株】自動生成：配当獲得履歴」の情報取得
  */
  const sheetDividendHist = spreadsheet.getSheetByName("【株】自動生成：配当獲得履歴");
  const lastRowNumDividendHist = sheetDividendHist.getLastRow();
  // ヘッダの取得
  //// カラム名が空白またはスペースのみになっているカラム以降は除外
  let headerDividendHist = sheetDividendHist.getRange("1:1").getValues()[0];
  headerDividendHist = headerDividendHist.slice(0, headerDividendHist.findIndex(v => v.match(REGEX_BLANK)));
  // console.log(headerDividendHist);
  // 処理に必要なカラム名のヘッダにおけるインデックスを取得
  const mapIdxColDividendHist = {
    "コード\n/シンボル": headerDividendHist.findIndex(v => v === "コード\n/シンボル"),
    "銘柄名": headerDividendHist.findIndex(v => v === "銘柄名"),
    "銘柄カテゴリ": headerDividendHist.findIndex(v => v === "銘柄カテゴリ"),
    "【配当取得】\n元URL": headerDividendHist.findIndex(v => v === "【配当取得】\n元URL"),
    "【配当取得】\nCSSSelector": headerDividendHist.findIndex(v => v === "【配当取得】\nCSSSelector"),
    "現在保有している株の保有開始日": headerDividendHist.findIndex(v => v === "現在保有している株の保有開始日"),
    "保有株数": headerDividendHist.findIndex(v => v === "保有株数"),
    "平均株価": headerDividendHist.findIndex(v => v === "平均株価"),
    "配当権利落ち日": headerDividendHist.findIndex(v => v === "配当権利落ち日"),
    "配当タイプ": headerDividendHist.findIndex(v => v === "配当タイプ"),
    "配当": headerDividendHist.findIndex(v => v === "配当"),
    "配当利回り": headerDividendHist.findIndex(v => v === "配当利回り"),
    "配当受渡総額": headerDividendHist.findIndex(v => v === "配当受渡総額"),
    "データ追加日": headerDividendHist.findIndex(v => v === "データ追加日"),
  };
  // console.log(mapIdxColDividendHist);

  /*
  シート「【株】自動生成：最新保有株」のデータをもとに配当獲得履歴を取得し，必要あればシート「【株】自動生成：配当獲得履歴」に追記書き込み．
  */
  // シート「【株】自動生成：最新保有株」のデータの読み込み
  //// ヘッダを除く2行目以降を取得
  const rowsStocksHolding = sheetStocksHolding
    .getRange(2, 1, lastRowNumStocksHolding, headerStocksHolding.length)
    .getValues();
  // console.log(rowsStocksHolding.slice(0, 10));
  // console.log(rowsStocksHolding.length);
  // コード/シンボルの一意な一覧取得
  const simbolsStocksHolding = [
    ...new Set(rowsStocksHolding.map(r => r[mapIdxColStocksHolding["コード\n/シンボル"]]))
  ].filter(v => !String(v).match(REGEX_BLANK));
  // console.log(simbolsStocksHolding);
  let rowsLatestDividend = [];
  return await Promise.all(
    simbolsStocksHolding.map(symbol =>
      insertRowLatestDividend(rowsLatestDividend, rowsStocksHolding, mapIdxColStocksHolding, mapIdxColDividendHist, symbol, TODAY)
    )
  )
    .then(_ => {
      // console.log(rowsLatestDividend); // この変数にシート「【株】自動生成：配当獲得履歴」に追加する，カラム順も整ったデータができているはず．
      /*
      シート「【株】自動生成：配当獲得履歴」にデータ追加
      */
      // 同一の配当権利落ち日・シンボルの組み合わせがあるデータをシートに追加しないよう除外しておく
      const rowsDividendHist = sheetDividendHist
        .getRange(2, 1, lastRowNumDividendHist, headerDividendHist.length)
        .getValues();
      const registeredSymbolAndExRightDate = new Set(rowsDividendHist.map(r => `${r[mapIdxColDividendHist["コード\n/シンボル"]]}_${getYyyymmdd(r[mapIdxColDividendHist["配当権利落ち日"]])}`));
      // console.log(registeredSymbolAndExRightDate.size);
      rowsLatestDividend = rowsLatestDividend.filter(r => !registeredSymbolAndExRightDate.has(`${r[mapIdxColDividendHist["コード\n/シンボル"]]}_${getYyyymmdd(r[mapIdxColDividendHist["配当権利落ち日"]])}`));
      // console.log(rowsLatestDividend);
      if (rowsLatestDividend.length > 0) {
        // 最後尾の次の行以降にデータを書き込み
        // console.log(lastRowNumDividendHist + 1, 1, rowsLatestDividend.length, Object.keys(mapIdxColDividendHist).length);
        sheetDividendHist.getRange(lastRowNumDividendHist + 1, 1, rowsLatestDividend.length, Object.keys(mapIdxColDividendHist).length)
          .setValues(rowsLatestDividend.sort(sortByExRightDateAndSymbolAsc(mapIdxColDividendHist["配当権利落ち日"], mapIdxColDividendHist["コード\n/シンボル"])));
      }
    })
    .catch(err => {
      throw err;
    });
}

async function insertRowLatestDividend(ret, rowsStocksHolding, mapIdxColStocksHolding, mapIdxColDividendHist, symbol, TODAY) {
  /*
  シンボルごとに直近の配当獲得履歴を導出して，直近配当獲得履歴リストに登録する．

  - Args
    - ret:Array: 処理結果のinsert先の配列
    - rowsStocksHolding:Array: 取引履歴データ
    - mapIdxColStocksHolding:Array: シート「【株】自動生成：最新保有株」のヘッダのカラム名とインデックス対応表
    - mapIdxColDividendHist:Array: シート「【株】自動生成：配当獲得履歴」のヘッダのカラム名とインデックス対応表
    - symbol:string|number: コード/シンボル
  */
  // あるシンボルの取引履歴を抽出
  //// シート「【株】自動生成：最新保有株」ではシンボルごとに一行登録されているはず．
  const rows = rowsStocksHolding
    .filter(r => symbol === r[mapIdxColStocksHolding["コード\n/シンボル"]]);
  // console.info(rows);
  if (rows.length !== 1) {
    throw Error(`{"コード\n/シンボル": "${symbol}"}をシート「【株】自動生成：最新保有株」から一意に取得することができませんでした．`);
  }
  const latestDividendsInfo = scrapeLatestDividendsInfoFromInvestingcom(rows[0], mapIdxColStocksHolding);
  // console.info(latestDividendsInfo);

  // 先に辞書形式でカラム名とデータを紐づけインデックスでカラム順整えた後で，insertする．
  const latestDividendsData = latestDividendsInfo.map(latestDividendInfo => ({
    "コード\n/シンボル": symbol,
    "銘柄名": rows[0][mapIdxColStocksHolding["銘柄名"]],
    "銘柄カテゴリ": rows[0][mapIdxColStocksHolding["銘柄カテゴリ"]],
    "【配当取得】\n元URL": rows[0][mapIdxColStocksHolding["【配当取得】\n元URL"]],
    "【配当取得】\nCSSSelector": rows[0][mapIdxColStocksHolding["【配当取得】\nCSSSelector"]],
    "現在保有している株の保有開始日": rows[0][mapIdxColStocksHolding["現在保有している株の保有開始日"]],
    "保有株数": rows[0][mapIdxColStocksHolding["保有株数"]],
    "平均株価": rows[0][mapIdxColStocksHolding["平均株価"]],
    "配当権利落ち日": latestDividendInfo["権利落ち日"],
    "配当タイプ": latestDividendInfo["タイプ"],
    "配当": latestDividendInfo["配当"],
    "配当利回り": latestDividendInfo["利回り"],
    "配当受渡総額": latestDividendInfo["配当"] * rows[0][mapIdxColStocksHolding["保有株数"]],
    "データ追加日": getYyyymmdd(TODAY),
  }));
  // console.info(latestDividendsData);

  let rowsLatestDividend = [];
  for (latestDividend of latestDividendsData) {
    let rowLatestDividend = [];
    try {
      for (let i = 0; i < Object.keys(mapIdxColDividendHist).length; i++) {
        const colNames = Object.keys(mapIdxColDividendHist).filter(k => mapIdxColDividendHist[k] === i);
        // console.info(colNames);
        if (colNames.length === 0) {
          throw Error("`mapIdxColDividendHist`の宣言時における，シート「【株】自動生成：配当獲得履歴」のカラム名の取得が十分にできていない可能性があります．");
        } else if (colNames.length > 1) {
          throw Error("`mapIdxColDividendHist`の宣言時における，シート「【株】自動生成：配当獲得履歴」のカラム名の取得が適切でない可能性があります．")
        }
        rowLatestDividend.push(latestDividend[colNames[0]]);
      }
    } catch (err) {
      throw err;
    }
    rowsLatestDividend.push(rowLatestDividend);
  }
  // console.info(rowsLatestDividend);
  // return rowsLatestDividend;
  ret.push(...rowsLatestDividend);
}

function scrapeLatestDividendsInfoFromInvestingcom(rowStockHolding, mapIdxColStocksHolding) {
  const res = UrlFetchApp.fetch(rowStockHolding[mapIdxColStocksHolding["【配当取得】\n元URL"]]);
  const strHtml = res.getContentText("UTF-8");
  // console.info(strHtml);
  const html = HtmlParser.parse(strHtml);
  // console.info(html);
  const table = html.querySelectorAll(rowStockHolding[mapIdxColStocksHolding["【配当取得】\nCSSSelector"]]);
  // console.info(table);
  if (table.length !== 1) {
    throw Error("tableタグを一意にスクレイピングできませんでした．");
  }
  // const colNames = table[0].querySelectorAll("thead>tr:nth-child(1)>th");
  // investing.comの配当ページから直近5行を取得して，さらに現在保有している株の保有開始日以降の配当情報を抽出．
  //// WARNING: 中間/期末/ボーナスといった配当のタイプによって複数抽出される場合があることに注意．
  const trs = table[0].querySelectorAll("tbody>tr:nth-child(-n+5)");
  let rowsCurrDividend = [];
  for (const tr of trs) {
    const tds = tr.querySelectorAll("td");
    const exRightDate = new Date(tds[0].innerText.replace(/(年|月)/g, "/").replace("日", ""));
    if (exRightDate >= rowStockHolding[mapIdxColStocksHolding["現在保有している株の保有開始日"]]) {
      // 現在保有している株の保有開始日以降の配当権利落ち日のみ抽出
      rowsCurrDividend.push({
        "権利落ち日": exRightDate,
        "配当": parseInt(tds[1].innerText),
        "タイプ": tds[2].querySelector("span").getAttribute("title"),
        "支払開始日": tds[3].innerText,
        "利回り": parseFloat(tds[4].innerText.replace("%", "")),
      });
    }
  }
  // console.info(rowsCurrDividend);
  if (rowsCurrDividend.length === 0) {
    return [];
  }
  const latestExRightDate = rowsCurrDividend.reduce((accumulator, r) => (accumulator["権利落ち日"] > r["権利落ち日"]) ? accumulator : r)["権利落ち日"];
  // console.info(latestExRightDate);
  // WARNING: year/month/dayでANDして等値判定すべきであり，そうしている
  // MEMO: 直近の配当権利落ち日より前も含めて取得したい場合，下記のfilterをコメントアウトすればよい．
  return rowsCurrDividend.filter(r => getYyyymmdd(r["権利落ち日"]) === getYyyymmdd(latestExRightDate));
}

function getYyyymmdd(_dt, pattern = "yyyy/mm/dd") {
  /*
  0埋めされた"yyyy"(year)・"mm"(month)・"dd"(date)をpattern通りに配置して返す．
  */
  let dt = null;
  if (typeOf(_dt, Date)) {
    dt = _dt;
  } else {
    dt = new Date(_dt);
  }
  return pattern
    .replace("yyyy", forwardPadding("0", 4, dt.getFullYear().toString())) // ("0000" + dt.getFullYear()).slice(-4)
    .replace("mm", forwardPadding("0", 2, (dt.getMonth() + 1).toString())) // ("00" + (dt.getMonth() + 1)).slice(-2)
    .replace("dd", forwardPadding("0", 2, dt.getDate().toString())); // ("00" + dt.getDate()).slice(-2)
}

function forwardPadding(padding, length, text) {
  /*
  指定文字数のうち前方を指定文字で埋めて返す．
  - Args
    - padding:String: 埋める文字
    - lengh:int: 文字数
    - text:String: 右詰めする文字列
  */
  length = length >= 0 ? length : -length;
  return  (String(padding).repeat(length) + text).slice(-length);
}

function typeOf(ins, type) {
  /*
  ins(instance)の型がtypeにマッチするか．
  - Refs
    - https://qiita.com/south37/items/c8d20a069fcbfe4fce85#constructor%E3%83%97%E3%83%AD%E3%83%91%E3%83%86%E3%82%A3%E3%82%92%E7%94%A8%E3%81%84%E3%81%9F%E5%88%A4%E5%AE%9A
  */
  return ins.constructor === type;
}

function sortByExRightDateAndSymbolAsc(idxColExRightDate, idxColSymbol) {
  /*
  「ORDER BY 配当権利落ち日, シンボル ASC」のソートに用いる．
  - Args
    - idxColExRightDate:number: ソートする比較対象のデータにおける配当権利落ち日のカラムインデックス
    - idxColSymbol:number: ソートする比較対象のデータにおけるシンボルのカラムインデックス
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
    if (!a[idxColExRightDate] || !a[idxColSymbol] || !b[idxColExRightDate] || !b[idxColSymbol]) {
      throw Error("「ORDER BY 配当権利落ち日, シンボル ASC」のソートするために，配当権利落ち日とシンボルが埋められている必要があります．");
    }
    const aTradeDate = new Date(a[idxColExRightDate]);
    const bTradeDate = new Date(b[idxColExRightDate]);
    return (aTradeDate - bTradeDate) || (String(a[idxColSymbol]).localeCompare(String(b[idxColSymbol])));
  }
  return validate;
}
