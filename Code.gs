/**
 * Night Walker / NW watcher 中継スクリプト（複数ペア対応版）
 *
 * 子の端末(Night Walker)から進捗を受け取ってスプレッドシートに記録し、
 * 親の端末(NW watcher)からの問い合わせに最新状況を返します。
 *
 * 「ペアコード」で親子を紐づけるため、1つのスプレッドシートを
 * 複数の親子が同時に使えます。
 *
 * 【設置手順】
 *  1. Googleスプレッドシートを新規作成
 *  2. 拡張機能 → Apps Script を開き、このコードを全て貼り付け
 *  3. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *       次のユーザーとして実行：自分
 *       アクセスできるユーザー：全員
 *  4. 発行されたウェブアプリURLを、Night Walker と NW watcher の両方に設定
 *
 * 【ペアコードについて】
 *  ・Night Walker の「ペアコードを作る」で自動生成されます（例：NW-8F3K2Q）
 *  ・同じコードを NW watcher に入力すると、その親子だけが紐づきます
 *  ・コードを知っている人だけが記録を見られるため、他人には教えないでください
 */

var SHEET_NAME = 'log';
var MAX_ROWS = 5000;      // これを超えたら古い行から削除する
var MIN_PAIR_LEN = 6;     // ペアコードの最低文字数

/** 子の端末から進捗を受け取る */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var pair = String(data.pair || '').trim();
    if (pair.length < MIN_PAIR_LEN) {
      return jsonOut({ ok: false, error: 'invalid pair code' });
    }
    var sh = getSheet();
    sh.appendRow([
      new Date(),
      pair,
      String(data.name || ''),
      String(data.planKey || ''),
      String(data.event || ''),
      (data.legId === undefined || data.legId === null) ? '' : data.legId,
      (data.lat === undefined || data.lat === null) ? '' : data.lat,
      (data.lng === undefined || data.lng === null) ? '' : data.lng,
      JSON.stringify(data.actuals || {}),
      JSON.stringify(data.weather || {}),
      JSON.stringify(data.route || null)
    ]);
    trimOldRows(sh);
    return jsonOut({ ok: true });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/** 親の端末へ最新状況を返す／記録をリセットする */
function doGet(e) {
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var pair = String(params.pair || '').trim();
    if (pair.length < MIN_PAIR_LEN) {
      return jsonOut({ ok: false, error: 'invalid pair code' });
    }

    var sh = getSheet();

    // 記録のリセット（このペアの記録だけを消す）
    if (params.action === 'reset') {
      var deleted = deleteRowsForPair(sh, pair);
      return jsonOut({ ok: true, reset: true, deleted: deleted });
    }

    var row = findLatestRowForPair(sh, pair);
    if (!row) {
      return jsonOut({ ok: true, empty: true });
    }
    var actuals = {}, weather = {}, routeCfg = null;
    try { actuals = row[8] ? JSON.parse(row[8]) : {}; } catch (e2) { actuals = {}; }
    try { weather = row[9] ? JSON.parse(row[9]) : {}; } catch (e3) { weather = {}; }
    try { routeCfg = row[10] ? JSON.parse(row[10]) : null; } catch (e4) { routeCfg = null; }

    return jsonOut({
      ok: true,
      empty: false,
      updatedAt: (row[0] instanceof Date) ? row[0].toISOString() : String(row[0]),
      pair: row[1],
      name: row[2],
      planKey: row[3],
      event: row[4],
      legId: row[5],
      lat: row[6] === '' ? null : Number(row[6]),
      lng: row[7] === '' ? null : Number(row[7]),
      actuals: actuals,
      weather: weather,
      route: routeCfg
    });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

/** 指定ペアの最新行を、下から探す */
function findLatestRowForPair(sh, pair) {
  var last = sh.getLastRow();
  if (last < 2) return null;
  var values = sh.getRange(2, 1, last - 1, 11).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][1]).trim() === pair) return values[i];
  }
  return null;
}

/** 指定ペアの行をすべて削除する */
function deleteRowsForPair(sh, pair) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var values = sh.getRange(2, 1, last - 1, 11).getValues();
  var count = 0;
  // 下から削除しないと行番号がずれる
  for (var i = values.length - 1; i >= 0; i--) {
    if (String(values[i][1]).trim() === pair) {
      sh.deleteRow(i + 2);
      count++;
    }
  }
  return count;
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['updatedAt', 'pair', 'name', 'planKey', 'event',
                  'legId', 'lat', 'lng', 'actuals', 'weather', 'route']);
  }
  return sh;
}

function trimOldRows(sh) {
  var last = sh.getLastRow();
  if (last > MAX_ROWS) {
    sh.deleteRows(2, last - MAX_ROWS);
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
