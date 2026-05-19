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
  MAX_RESULTS:             20,          // 最終的に結果シートに出力する件数
  PAGE_SIZE:               50,          // 1ページあたりの取得件数（API上限）
  MAX_PAGES:               5,           // 最大ページ数（クォータ節約のための安全装置）
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

  SpreadsheetApp.getUi().alert("✅ シートの初期化が完了しました！\n「検索」シートのB2セルにチャンネルを入力して、メニューから検索を実行してください。");
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

    // --- ⑤ アップロードプレイリストから全動画IDを収集（全期間・調査済み除外）---
    // Search APIは全期間の再生数順に弱いため、Uploads Playlistから全取得→自前ソートに切替
    var channelDetail = YouTube.Channels.list("contentDetails", { id: channelId });
    if (!channelDetail.items || channelDetail.items.length === 0) {
      ui.alert("❌ チャンネル情報の取得失敗", "チャンネルの詳細情報をYouTube APIから取得できませんでした。\nしばらく待ってから再試行してください。", ui.ButtonSet.OK);
      return;
    }
    var uploadsPlaylistId = channelDetail.items[0].contentDetails.relatedPlaylists.uploads;
    if (!uploadsPlaylistId) {
      ui.alert("❌ プレイリスト取得失敗", "アップロードプレイリストIDを取得できませんでした。\nチャンネルが正しいか確認してください。", ui.ButtonSet.OK);
      return;
    }
    Logger.log("アップロードプレイリストID: " + uploadsPlaylistId);

    var uninvestigatedIds = [];
    var skippedCount = 0;
    var playlistPageToken = null;
    var playlistPageCount = 0;

    while (playlistPageCount < CONFIG.MAX_PAGES) {
      var playlistParams = { playlistId: uploadsPlaylistId, maxResults: CONFIG.PAGE_SIZE };
      if (playlistPageToken) playlistParams.pageToken = playlistPageToken;

      var playlistResponse = YouTube.PlaylistItems.list("contentDetails", playlistParams);
      playlistPageCount++;

      if (!playlistResponse.items || playlistResponse.items.length === 0) break;

      playlistResponse.items.forEach(function(item) {
        var vid = item.contentDetails.videoId;
        if (investigatedIds.has(vid)) {
          skippedCount++;
        } else {
          uninvestigatedIds.push(vid);
        }
      });

      Logger.log("プレイリストページ" + playlistPageCount + ": 未調査累計=" + uninvestigatedIds.length + "件");
      playlistPageToken = playlistResponse.nextPageToken;
      if (!playlistPageToken) break;
    }

    if (uninvestigatedIds.length === 0) {
      ui.alert("ℹ️ 新規動画なし", "「" + channelName + "」の未調査動画が見つかりませんでした。\n（調査済み: " + investigatedIds.size + " 件）", ui.ButtonSet.OK);
      return;
    }

    Logger.log("未調査動画ID: " + uninvestigatedIds.length + "件");

    // --- ⑥ 未調査動画の統計情報を50件ずつバッチ取得 ---
    var allVideos = [];
    for (var i = 0; i < uninvestigatedIds.length; i += CONFIG.PAGE_SIZE) {
      var batchIds = uninvestigatedIds.slice(i, i + CONFIG.PAGE_SIZE).join(",");
      var batchResponse = YouTube.Videos.list("snippet,statistics", { id: batchIds, maxResults: CONFIG.PAGE_SIZE });
      if (batchResponse.items) {
        batchResponse.items.forEach(function(item) {
          allVideos.push({
            videoId:      item.id,
            title:        item.snippet.title,
            url:          "https://www.youtube.com/watch?v=" + item.id,
            channelTitle: item.snippet.channelTitle,
            publishedAt:  item.snippet.publishedAt,
            viewCount:    parseInt(item.statistics.viewCount || "0", 10)
          });
        });
      }
    }

    // --- ⑦ 再生数で降順ソートして上位MAX_RESULTS件を選ぶ ---
    allVideos.sort(function(a, b) { return b.viewCount - a.viewCount; });
    var videos = allVideos.slice(0, CONFIG.MAX_RESULTS);

    Logger.log("全期間再生数ソート後、上位" + videos.length + "件を出力");

    // --- ⑧ 結果シートへの書き出し ---
    var rows = videos.map(function(video, index) {
      return [
        index + 1,
        video.title,
        video.url,
        video.viewCount,
        formatDate_(video.publishedAt),
        video.channelTitle
      ];
    });

    if (rows.length > 0) {
      resultSheet.getRange(CONFIG.RESULT_START_ROW, 1, rows.length, 6).setValues(rows);
      resultSheet.getRange(CONFIG.RESULT_START_ROW, 4, rows.length, 1).setNumberFormat("#,##0");
      applyAlternateRowColors_(resultSheet, CONFIG.RESULT_START_ROW, rows.length);
      recordSearchMeta_(resultSheet, channelName, rows.length, skippedCount, uninvestigatedIds.length);
    }

    // --- ⑨ 調査済みシートに今回の動画を追記 ---
    appendToInvestigated_(videos);

    ss.setActiveSheet(resultSheet);

    ui.alert(
      "✅ 検索完了！",
      "「" + channelName + "」の新規動画 " + rows.length + " 件をシートに出力しました。\n" +
      "（調査済み除外: " + skippedCount + " 件 / 未調査プール: " + uninvestigatedIds.length + " 件）",
      ui.ButtonSet.OK
    );

  } catch (e) {
    Logger.log("エラー発生: " + e.toString());
    var errorMsg = e.toString();
    var title = "❌ エラーが発生しました";
    var body;

    if (errorMsg.indexOf("quota") !== -1 || errorMsg.indexOf("quotaExceeded") !== -1) {
      title = "❌ APIクォータエラー";
      body = "YouTube APIの1日あたりの呼び出し上限に達しました。\n明日以降に再試行してください。";
    } else if (errorMsg.indexOf("500") !== -1 || errorMsg.indexOf("503") !== -1 || errorMsg.indexOf("Backend") !== -1) {
      body = "YouTube APIサーバーが一時的に不安定です。\n数分待ってから再試行してください。";
    } else if (errorMsg.indexOf("TypeError") !== -1) {
      body = "データの取得中に予期しないエラーが発生しました。\n再試行してください。それでも続く場合はチャンネルIDを確認してください。";
    } else {
      body = "処理中にエラーが発生しました。\nYouTube Data API v3 が有効になっているか確認してください。";
    }

    ui.alert(title, body + "\n\n【エラー詳細】\n" + errorMsg, ui.ButtonSet.OK);
  }
}

// ============================================================
// 調査済み管理
// ============================================================

/**
 * 調査済みシートを作成してヘッダーを設置する
 * 列構成: 動画ID | 動画タイトル | 動画URL | チャンネル名 | 公開日 | 再生回数 | 調査日
 */
function initializeInvestigatedSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.INVESTIGATED_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.INVESTIGATED_SHEET_NAME);
  }

  var headers = ["動画ID", "動画タイトル", "動画URL", "チャンネル名", "公開日", "再生回数", "調査日"];
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers])
    .setBackground("#e65100")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.setColumnWidth(1, 160);  // 動画ID
  sheet.setColumnWidth(2, 380);  // 動画タイトル
  sheet.setColumnWidth(3, 300);  // 動画URL
  sheet.setColumnWidth(4, 200);  // チャンネル名
  sheet.setColumnWidth(5, 100);  // 公開日
  sheet.setColumnWidth(6, 120);  // 再生回数
  sheet.setColumnWidth(7, 140);  // 調査日
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
 * @param {Array} videos - {videoId, title, url, channelTitle, publishedAt, viewCount}の配列
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
      video.url,
      video.channelTitle,
      formatDate_(video.publishedAt),
      video.viewCount,
      today
    ];
  });

  var lastRow = sheet.getLastRow();
  var startRow = Math.max(lastRow + 1, 2);
  sheet.getRange(startRow, 1, rows.length, 7).setValues(rows);
  sheet.getRange(startRow, 6, rows.length, 1).setNumberFormat("#,##0");  // 再生回数（F列）

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
function recordSearchMeta_(sheet, keyword, count, skipped, pool) {
  sheet.getRange("H1").setValue("チャンネル名:").setFontWeight("bold");
  sheet.getRange("I1").setValue(keyword).setBackground("#fff9c4");

  sheet.getRange("H2").setValue("実行日時:").setFontWeight("bold");
  var now = new Date();
  var formattedNow = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");
  sheet.getRange("I2").setValue(formattedNow);

  sheet.getRange("H3").setValue("取得件数:").setFontWeight("bold");
  sheet.getRange("I3").setValue(count + " 件");

  sheet.getRange("H4").setValue("除外件数:").setFontWeight("bold");
  sheet.getRange("I4").setValue(skipped + " 件（調査済み）");

  sheet.getRange("H5").setValue("未調査プール:").setFontWeight("bold");
  sheet.getRange("I5").setValue(pool + " 件");

  sheet.setColumnWidth(8, 140);
  sheet.setColumnWidth(9, 220);
}
