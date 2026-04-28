// ============================================================
// YouTubeチャンネル動画抽出ツール
// Google Apps Script + YouTube Data API v3
// ============================================================

// ---- 設定値 (必要に応じて変更してください) ----
var CONFIG = {
  SEARCH_SHEET_NAME:       "検索",      // チャンネル入力・ボタン配置シート名
  RESULT_SHEET_NAME:       "結果",      // 検索結果を出力するシート名
  INVESTIGATED_SHEET_NAME: "調査済み", // 調査済み動画を蓄積するシート名
  CHANNEL_CELL:            "B2",        // チャンネルID/URLを入力するセル番地
  MAX_RESULTS:             10,          // 最終的に結果シートに出力する件数
  FETCH_BUFFER_SIZE:       50,          // APIから一度に取得する件数（調査済み除外のバッファ）
  RESULT_START_ROW:        2,           // 結果を書き出す開始行（1行目はヘッダー）
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
  searchSheet.getRange("B2").setValue("").setBackground("#fff9c4");
  searchSheet.getRange("A3").setValue("例: UCxxxxxx / @handle / チャンネルURL").setFontColor("#999999").setFontSize(9);

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

  // --- 調査済みシートの準備 ---
  initializeInvestigatedSheet_();

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

    var channelInfo = YouTube.Channels.list("snippet", { id: channelId });
    var channelName = (channelInfo.items && channelInfo.items.length > 0)
      ? channelInfo.items[0].snippet.title
      : channelId;

    Logger.log("チャンネルID: " + channelId + " / チャンネル名: " + channelName);

    // --- ③ 調査済み動画IDのセットを取得 ---
    var investigatedIds = getInvestigatedVideoIds_();
    Logger.log("調査済み件数: " + investigatedIds.size);

    // --- ④ 結果シートの取得と前回データのクリア ---
    var resultSheet = ss.getSheetByName(CONFIG.RESULT_SHEET_NAME);
    if (!resultSheet) {
      resultSheet = ss.insertSheet(CONFIG.RESULT_SHEET_NAME);
    }
    clearResultData_(resultSheet);
    setResultHeader_(resultSheet);

    // --- ⑤ 「ぴったり1ヶ月前」の日付を計算する ---
    var publishedAfter = getOneMonthAgoISOString_();
    Logger.log("検索期間: " + publishedAfter + " 以降");

    // --- ⑥ YouTube Search API: バッファサイズ分だけ取得して調査済みを除外 ---
    var searchResponse = YouTube.Search.list("id,snippet", {
      channelId:      channelId,
      type:           "video",
      order:          "viewCount",
      publishedAfter: publishedAfter,
      maxResults:     CONFIG.FETCH_BUFFER_SIZE,
    });

    if (!searchResponse.items || searchResponse.items.length === 0) {
      ui.alert("ℹ️ 検索結果なし", "「" + channelName + "」の過去1ヶ月以内の動画が見つかりませんでした。", ui.ButtonSet.OK);
      return;
    }

    // 調査済みを除外してから最大MAX_RESULTS件に絞る
    var newItems = searchResponse.items.filter(function(item) {
      return !investigatedIds.has(item.id.videoId);
    });

    if (newItems.length === 0) {
      ui.alert("ℹ️ 新規動画なし", "「" + channelName + "」の過去1ヶ月以内の動画はすべて調査済みです。", ui.ButtonSet.OK);
      return;
    }

    var targetItems = newItems.slice(0, CONFIG.MAX_RESULTS);
    var videoIds = targetItems.map(function(item) { return item.id.videoId; }).join(",");

    Logger.log("新規videoId一覧 (" + targetItems.length + "件): " + videoIds);

    // --- ⑦ YouTube Videos API: 統計情報を取得 ---
    var videosResponse = YouTube.Videos.list("snippet,statistics", { id: videoIds });

    if (!videosResponse.items || videosResponse.items.length === 0) {
      ui.alert("❌ エラー", "動画の詳細情報の取得に失敗しました。しばらく時間をおいて再試行してください。", ui.ButtonSet.OK);
      return;
    }

    // --- ⑧ 取得したデータを再生数で降順ソート ---
    var videos = videosResponse.items.map(function(item) {
      return {
        videoId:      item.id,
        title:        item.snippet.title,
        channelTitle: item.snippet.channelTitle,
        publishedAt:  item.snippet.publishedAt,
        viewCount:    parseInt(item.statistics.viewCount || "0", 10)
      };
    });

    videos.sort(function(a, b) { return b.viewCount - a.viewCount; });

    // --- ⑨ 結果シートへの書き出し ---
    var rows = videos.map(function(video, index) {
      var videoUrl = "https://www.youtube.com/watch?v=" + video.videoId;
      return [
        index + 1,
        video.title,
        videoUrl,
        video.viewCount,
        formatDate_(video.publishedAt),
        video.channelTitle
      ];
    });

    if (rows.length > 0) {
      resultSheet.getRange(CONFIG.RESULT_START_ROW, 1, rows.length, 6).setValues(rows);
      resultSheet.getRange(CONFIG.RESULT_START_ROW, 4, rows.length, 1).setNumberFormat("#,##0");
      applyAlternateRowColors_(resultSheet, CONFIG.RESULT_START_ROW, rows.length);
      recordSearchMeta_(resultSheet, channelName, rows.length);
    }

    // --- ⑩ 調査済みシートに今回の動画を追記 ---
    appendToInvestigated_(videos);

    ss.setActiveSheet(resultSheet);

    ui.alert(
      "✅ 検索完了！",
      "「" + channelName + "」の新規動画 " + rows.length + " 件をシートに出力しました。\n" +
      "（調査済み除外: " + (searchResponse.items.length - newItems.length) + " 件）",
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log("エラー発生: " + e.toString());
    var errorMsg = e.toString();
    if (errorMsg.indexOf("quota") !== -1 || errorMsg.indexOf("quotaExceeded") !== -1) {
      ui.alert("❌ APIクォータエラー", "YouTube APIの1日あたりの呼び出し上限に達しました。\n明日以降に再試行してください。\n\nエラー詳細: " + errorMsg, ui.ButtonSet.OK);
    } else {
      ui.alert("❌ エラーが発生しました", "処理中にエラーが発生しました。\nAPIの設定（YouTube Data API v3が有効か）を確認してください。\n\nエラー詳細: " + errorMsg, ui.ButtonSet.OK);
    }
  }
}

// ============================================================
// 調査済み管理
// ============================================================

/**
 * 調査済みシートを作成してヘッダーを設置する
 */
function initializeInvestigatedSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INVESTIGATED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.INVESTIGATED_SHEET_NAME);
  }

  var headers = ["動画ID", "動画タイトル", "チャンネル名", "公開日", "再生回数", "調査日"];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground("#e65100")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 160);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 140);
  sheet.setFrozenRows(1);

  return sheet;
}

/**
 * 調査済みシートからすでに調査した動画IDのSetを返す
 */
function getInvestigatedVideoIds_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INVESTIGATED_SHEET_NAME);
  var ids = new Set();

  if (!sheet) return ids;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return ids;

  // A列（動画ID）を一括取得
  var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  values.forEach(function(row) {
    if (row[0]) ids.add(row[0].toString());
  });

  return ids;
}

/**
 * 調査した動画を調査済みシートの末尾に追記する
 * @param {Array} videos - {videoId, title, channelTitle, publishedAt, viewCount}の配列
 */
function appendToInvestigated_(videos) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INVESTIGATED_SHEET_NAME);
  if (!sheet) {
    sheet = initializeInvestigatedSheet_();
  }

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy/MM/dd");
  var rows = videos.map(function(video) {
    return [
      video.videoId,
      video.title,
      video.channelTitle,
      formatDate_(video.publishedAt),
      video.viewCount,
      today
    ];
  });

  var lastRow = sheet.getLastRow();
  var startRow = Math.max(lastRow + 1, 2);
  sheet.getRange(startRow, 1, rows.length, 6).setValues(rows);
  sheet.getRange(startRow, 5, rows.length, 1).setNumberFormat("#,##0");

  Logger.log("調査済みシートに " + rows.length + " 件追記しました（合計: " + (startRow - 2 + rows.length) + " 件）");
}

// ============================================================
// ヘルパー関数群
// ============================================================

/**
 * チャンネルID・ハンドル・URLからチャンネルIDを解決して返す
 */
function resolveChannelId_(input) {
  var channelMatch = input.match(/youtube\.com\/channel\/(UC[\w-]+)/);
  if (channelMatch) return channelMatch[1];

  if (/^UC[\w-]+$/.test(input)) return input;

  var handleMatch = input.match(/(?:youtube\.com\/)?@([^\s/?&]+)/);
  var handle = handleMatch ? handleMatch[1] : null;

  if (!handle) {
    var customMatch = input.match(/youtube\.com\/(?:c|user)\/([^\s/?&]+)/);
    if (customMatch) handle = customMatch[1];
  }

  if (handle) {
    var res = YouTube.Channels.list("id", { forHandle: handle });
    if (res.items && res.items.length > 0) return res.items[0].id;
  }

  return null;
}

/**
 * 本日から「ぴったり1ヶ月前」のISO 8601形式の文字列を返す
 */
function getOneMonthAgoISOString_() {
  var now = new Date();
  var oneMonthAgo = new Date(now.getFullYear(), now.getMonth() - 1, now.getDate(), 0, 0, 0);
  return oneMonthAgo.toISOString();
}

/**
 * ISO 8601形式の日付文字列を "YYYY/MM/DD" 形式に変換する
 */
function formatDate_(isoString) {
  if (!isoString) return "";
  var date = new Date(isoString);
  var y = date.getFullYear();
  var m = ("0" + (date.getMonth() + 1)).slice(-2);
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
    .setBackground("#34a853")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 400);
  sheet.setColumnWidth(3, 320);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 200);
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
}

/**
 * 結果行に交互の背景色を設定する
 */
function applyAlternateRowColors_(sheet, startRow, numRows) {
  for (var i = 0; i < numRows; i++) {
    var row = startRow + i;
    var color = (i % 2 === 0) ? "#ffffff" : "#f1f8e9";
    sheet.getRange(row, 1, 1, 6).setBackground(color);
  }
}

/**
 * 検索チャンネル名と実行日時を結果シートの右上に記録する
 */
function recordSearchMeta_(sheet, keyword, count) {
  sheet.getRange("H1").setValue("チャンネル名:").setFontWeight("bold");
  sheet.getRange("I1").setValue(keyword).setBackground("#fff9c4");

  sheet.getRange("H2").setValue("実行日時:").setFontWeight("bold");
  var now = new Date();
  var formattedNow = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  sheet.getRange("I2").setValue(formattedNow);

  sheet.getRange("H3").setValue("取得件数:").setFontWeight("bold");
  sheet.getRange("I3").setValue(count + " 件");

  sheet.setColumnWidth(8, 140);
  sheet.setColumnWidth(9, 200);
}
