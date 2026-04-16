// ============================================================
// YouTubeチャンネル動画抽出ツール
// Google Apps Script + YouTube Data API v3
// ============================================================

// ---- 設定値 (必要に応じて変更してください) ----
var CONFIG = {
  SEARCH_SHEET_NAME: "検索",          // チャンネル入力・ボタン配置シート名
  RESULT_SHEET_NAME: "結果",          // 検索結果を出力するシート名
  CHANNEL_CELL: "B2",                 // チャンネルID/URLを入力するセル番地
  MAX_RESULTS: 10,                    // 取得する動画の最大件数
  RESULT_START_ROW: 2,                // 結果を書き出す開始行（1行目はヘッダー）
};

// ============================================================
// スプレッドシートを開いたときに自動でカスタムメニューを追加する
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("▶️ YouTube検索")
    .addItem("チャンネル動画を取得", "fetchYouTubeTrends")
    .addSeparator()
    .addItem("シートを初期化（ヘッダー設置）", "initializeSheets")
    .addToUi();
}

// ============================================================
// シートの初期化: シート作成 & ヘッダー行を設置する
// ============================================================
function initializeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 検索シートの準備 ---
  var searchSheet = ss.getSheetByName(CONFIG.SEARCH_SHEET_NAME);
  if (!searchSheet) {
    searchSheet = ss.insertSheet(CONFIG.SEARCH_SHEET_NAME);
  }
  searchSheet.getRange("A2").setValue("チャンネルID/URL:");
  searchSheet.getRange("B2").setValue("").setBackground("#fff9c4"); // 入力欄を黄色にして分かりやすく
  searchSheet.getRange("A3").setValue("例: UCxxxxxx / @handle / チャンネルURL").setFontColor("#999999").setFontSize(9);

  // A1:B1をヘッダー風にスタイリング
  searchSheet.getRange("A1:B1").merge()
    .setValue("🔍 YouTubeチャンネル動画検索ツール")
    .setFontSize(14)
    .setFontWeight("bold")
    .setBackground("#4285f4")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  searchSheet.setColumnWidth(1, 160);
  searchSheet.setColumnWidth(2, 300);

  // --- 結果シートの準備 ---
  var resultSheet = ss.getSheetByName(CONFIG.RESULT_SHEET_NAME);
  if (!resultSheet) {
    resultSheet = ss.insertSheet(CONFIG.RESULT_SHEET_NAME);
  }
  setResultHeader_(resultSheet);

  SpreadsheetApp.getUi().alert("✅ シートの初期化が完了しました！\n「検索」シートのB2セルにキーワードを入力して、メニューから検索を実行してください。");
}

// ============================================================
// メイン処理: YouTubeトレンド動画を取得してシートに書き出す
// ============================================================
function fetchYouTubeTrends() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    // --- ① チャンネルID/URLの取得とバリデーション ---
    var searchSheet = ss.getSheetByName(CONFIG.SEARCH_SHEET_NAME);
    if (!searchSheet) {
      ui.alert("❌ エラー", "「" + CONFIG.SEARCH_SHEET_NAME + "」シートが見つかりません。\nメニューから「シートを初期化」を先に実行してください。", ui.ButtonSet.OK);
      return;
    }

    var channelInput = searchSheet.getRange(CONFIG.CHANNEL_CELL).getValue().toString().trim();

    // 入力が空欄の場合は処理を中断
    if (!channelInput) {
      ui.alert("⚠️ チャンネルが未入力です", "「" + CONFIG.SEARCH_SHEET_NAME + "」シートの " + CONFIG.CHANNEL_CELL + " セルにチャンネルID・ハンドル・URLを入力してから実行してください。", ui.ButtonSet.OK);
      return;
    }

    // --- ② チャンネルIDを解決する ---
    var channelId = resolveChannelId_(channelInput);
    if (!channelId) {
      ui.alert("❌ チャンネルが見つかりません", "入力値からチャンネルを特定できませんでした。\nチャンネルID（UCxxxxxx）、ハンドル（@handle）、またはチャンネルURLを入力してください。", ui.ButtonSet.OK);
      return;
    }

    // チャンネル名を取得して表示用に保存
    var channelInfo = YouTube.Channels.list("snippet", { id: channelId });
    var channelName = (channelInfo.items && channelInfo.items.length > 0)
      ? channelInfo.items[0].snippet.title
      : channelId;

    Logger.log("チャンネルID: " + channelId + " / チャンネル名: " + channelName);

    // --- ③ 結果シートの取得と前回データのクリア ---
    var resultSheet = ss.getSheetByName(CONFIG.RESULT_SHEET_NAME);
    if (!resultSheet) {
      resultSheet = ss.insertSheet(CONFIG.RESULT_SHEET_NAME);
    }
    clearResultData_(resultSheet);
    setResultHeader_(resultSheet);

    // --- ④ 「ぴったり1ヶ月前」の日付を計算する ---
    var publishedAfter = getOneMonthAgoISOString_();
    Logger.log("検索期間: " + publishedAfter + " 以降");

    // --- ⑤ YouTube Data API: /search エンドポイントで動画IDを取得 ---
    var searchResponse = YouTube.Search.list("id,snippet", {
      channelId: channelId,            // 対象チャンネルを指定
      type: "video",                   // 動画のみ対象
      order: "viewCount",              // 再生回数順にソート
      publishedAfter: publishedAfter,  // 1ヶ月前以降に絞り込み
      maxResults: CONFIG.MAX_RESULTS,  // 最大10件
    });

    // 検索結果が0件の場合はアラートを出して終了
    if (!searchResponse.items || searchResponse.items.length === 0) {
      ui.alert("ℹ️ 検索結果なし", "「" + channelName + "」の過去1ヶ月以内の動画が見つかりませんでした。", ui.ButtonSet.OK);
      return;
    }

    // 取得した動画IDをカンマ区切りの文字列にまとめる
    var videoIds = searchResponse.items.map(function(item) {
      return item.id.videoId;
    }).join(",");

    Logger.log("取得したvideoId一覧: " + videoIds);

    // --- ⑤ YouTube Data API: /videos エンドポイントで正確な統計情報を取得 ---
    // /search では statistics が取れないため、videoId を使って再度クエリする
    var videosResponse = YouTube.Videos.list("snippet,statistics", {
      id: videoIds
    });

    if (!videosResponse.items || videosResponse.items.length === 0) {
      ui.alert("❌ エラー", "動画の詳細情報の取得に失敗しました。しばらく時間をおいて再試行してください。", ui.ButtonSet.OK);
      return;
    }

    // --- ⑥ 取得したデータを再生数で降順ソート ---
    var videos = videosResponse.items.map(function(item) {
      return {
        videoId: item.id,
        title: item.snippet.title,
        channelTitle: item.snippet.channelTitle,
        publishedAt: item.snippet.publishedAt,
        viewCount: parseInt(item.statistics.viewCount || "0", 10)
      };
    });

    // /videos APIの返却順は保証されないため、再生数で再ソートする
    videos.sort(function(a, b) {
      return b.viewCount - a.viewCount;
    });

    // --- ⑦ スプレッドシートへの書き出し ---
    var rows = videos.map(function(video, index) {
      var videoUrl = "https://www.youtube.com/watch?v=" + video.videoId;
      var publishedDate = formatDate_(video.publishedAt); // YYYY/MM/DD 形式に変換

      return [
        index + 1,           // 順位
        video.title,         // 動画タイトル
        videoUrl,            // 動画URL
        video.viewCount,     // 再生回数（数値）
        publishedDate,       // 公開日
        video.channelTitle   // チャンネル名
      ];
    });

    // データをまとめてシートに書き込む（1セルずつ書き込むより高速）
    if (rows.length > 0) {
      resultSheet.getRange(CONFIG.RESULT_START_ROW, 1, rows.length, 6).setValues(rows);

      // 再生回数列（D列）に数値フォーマットを適用して読みやすくする
      resultSheet.getRange(CONFIG.RESULT_START_ROW, 4, rows.length, 1).setNumberFormat("#,##0");

      // 行の交互背景色を設定して見やすくする
      applyAlternateRowColors_(resultSheet, CONFIG.RESULT_START_ROW, rows.length);

      // チャンネル名と実行日時を結果シートに記録する
      recordSearchMeta_(resultSheet, channelName, rows.length);
    }

    // アクティブシートを結果シートに切り替える
    ss.setActiveSheet(resultSheet);

    ui.alert("✅ 検索完了！", "「" + channelName + "」の検索結果 " + rows.length + " 件をシートに出力しました。", ui.ButtonSet.OK);

  } catch (e) {
    // APIクォータエラーや予期せぬエラーをキャッチしてユーザーに通知
    Logger.log("エラー発生: " + e.toString());
    var errorMsg = e.toString();

    // クォータエラーの場合はより分かりやすいメッセージを表示
    if (errorMsg.indexOf("quota") !== -1 || errorMsg.indexOf("quotaExceeded") !== -1) {
      ui.alert("❌ APIクォータエラー", "YouTube APIの1日あたりの呼び出し上限に達しました。\n明日以降に再試行してください。\n\nエラー詳細: " + errorMsg, ui.ButtonSet.OK);
    } else {
      ui.alert("❌ エラーが発生しました", "処理中にエラーが発生しました。\nAPIの設定（YouTube Data API v3が有効か）を確認してください。\n\nエラー詳細: " + errorMsg, ui.ButtonSet.OK);
    }
  }
}

// ============================================================
// ヘルパー関数群
// ============================================================

/**
 * チャンネルID・ハンドル・URLからチャンネルIDを解決して返す
 * - UCxxxxxx → そのまま返す
 * - @handle → channels.list API で解決
 * - https://www.youtube.com/channel/UCxxxxxx → IDを抽出
 * - https://www.youtube.com/@handle → ハンドルを抽出して解決
 * 解決できない場合は null を返す
 */
function resolveChannelId_(input) {
  // /channel/UCxxxxxx 形式のURLからIDを抽出
  var channelMatch = input.match(/youtube\.com\/channel\/(UC[\w-]+)/);
  if (channelMatch) return channelMatch[1];

  // UCxxxxxx 形式のIDはそのまま使用
  if (/^UC[\w-]+$/.test(input)) return input;

  // @handle または youtube.com/@handle からハンドルを抽出（日本語など Unicode 文字も対応）
  var handleMatch = input.match(/(?:youtube\.com\/)?@([^\s/?&]+)/);
  var handle = handleMatch ? handleMatch[1] : null;

  // ハンドルがない場合はカスタムURL（/c/name や /user/name）を試みる
  if (!handle) {
    var customMatch = input.match(/youtube\.com\/(?:c|user)\/([^\s/?&]+)/);
    if (customMatch) handle = customMatch[1];
  }

  if (handle) {
    // forHandle パラメータでチャンネルIDを検索
    var res = YouTube.Channels.list("id", { forHandle: handle });
    if (res.items && res.items.length > 0) return res.items[0].id;
  }

  return null;
}

/**
 * 本日から「ぴったり1ヶ月前」のISO 8601形式の文字列を返す
 * 例: 2024-06-08T00:00:00Z
 */
function getOneMonthAgoISOString_() {
  var now = new Date();
  // 1ヶ月前を計算（月をマイナス1する。月末の扱いはJavaScriptが自動調整）
  var oneMonthAgo = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate(), 0, 0, 0);
  return oneMonthAgo.toISOString();
}

/**
 * ISO 8601形式の日付文字列を "YYYY/MM/DD" 形式に変換する
 * 例: "2024-06-08T12:34:56Z" → "2024/06/08"
 */
function formatDate_(isoString) {
  if (!isoString) return "";
  var date = new Date(isoString);
  var y = date.getFullYear();
  var m = ("0" + (date.getMonth() + 1)).slice(-2); // 月は0始まりなので+1
  var d = ("0" + date.getDate()).slice(-2);
  return y + "/" + m + "/" + d;
}

/**
 * 結果シートのヘッダー行を設置・スタイリングする
 */
function setResultHeader_(sheet) {
  var headers = ["順位", "動画タイトル", "動画URL", "再生回数", "公開日", "チャンネル名"];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);

  headerRange.setValues([headers])
    .setBackground("#34a853")   // Googleグリーン
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // 列幅を内容に合わせて調整
  sheet.setColumnWidth(1, 60);   // 順位
  sheet.setColumnWidth(2, 400);  // 動画タイトル
  sheet.setColumnWidth(3, 320);  // 動画URL
  sheet.setColumnWidth(4, 120);  // 再生回数
  sheet.setColumnWidth(5, 100);  // 公開日
  sheet.setColumnWidth(6, 200);  // チャンネル名
  // 行を固定してスクロールしてもヘッダーが見えるようにする
  sheet.setFrozenRows(1);
}

/**
 * 結果シートの前回データ部分（ヘッダー以外）をクリアする
 */
function clearResultData_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow >= CONFIG.RESULT_START_ROW) {
    sheet.getRange(CONFIG.RESULT_START_ROW, 1, lastRow - CONFIG.RESULT_START_ROW + 1, 10).clearContent().clearFormat();
  }
  // シート上部のメタ情報エリアもクリア（後で上書きするため）
  // ヘッダー行(1行目)はsetResultHeader_で都度セットするのでここではクリアしない
}

/**
 * 結果行に交互の背景色（白・薄いグレー）を設定して可読性を高める
 */
function applyAlternateRowColors_(sheet, startRow, numRows) {
  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    var color = (i % 2 === 0) ? "#ffffff" : "#f1f8e9"; // 偶数行は白、奇数行は薄い緑
    sheet.getRange(row, 1, 1, 6).setBackground(color);
  }
}

/**
 * 検索キーワードと実行日時を結果シートの目立つ位置に記録する
 */
function recordSearchMeta_(sheet, keyword, count) {
  // シートの右上エリア（H1:J2）にメタ情報を表示
  sheet.getRange("H1").setValue("チャンネル名:").setFontWeight("bold");
  sheet.getRange("I1").setValue(keyword).setBackground("#fff9c4");

  sheet.getRange("H2").setValue("実行日時:").setFontWeight("bold");
  var now = new Date();
  var formattedNow = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  sheet.getRange("I2").setValue(formattedNow);

  sheet.getRange("H3").setValue("取得件数:").setFontWeight("bold");
  sheet.getRange("I3").setValue(count + " 件");

  sheet.setColumnWidth(8, 140);  // H列
  sheet.setColumnWidth(9, 200);  // I列
}
