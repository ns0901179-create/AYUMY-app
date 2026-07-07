/**
 * AYUMY SyncCal Pro v1.6 (Smart Sync & Delete Detection)
 */

const APP_ID = "AYUMY_SyncCal_pro_v1_3_0";
const LOGO_CACHE_KEY = "SYNC_CAL_LOGO_BASE64";

// ==========================================
// 1. 初期セットアップ
// ==========================================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return "データベース設定完了。祝日計算と画像処理を更新しました。";
}

// ==========================================
// 2. Web UI 表示用エンドポイント
// ==========================================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('AYUMY SyncCal 💜')
    // ここからPWA設定を追加します
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-status-bar-style', 'black-translucent')
    .addMetaTag('apple-mobile-web-app-title', 'AYUMY SyncCal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 3. 各種設定値取得ヘルパー
// ==========================================
function getConfigValue_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("CONFIG");
  if (!configSheet) return "";
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) return data[i][1] ? data[i][1].toString().trim() : "";
  }
  return "";
}

function getTargetCalendar_() {
  const calId = getConfigValue_("CALENDAR_ID");
  if (!calId || calId === "primary") return CalendarApp.getDefaultCalendar();
  const cal = CalendarApp.getCalendarById(calId);
  if (!cal) throw new Error("指定されたカレンダーIDが見つかりません。");
  return cal;
}

// ==========================================
// 4. データ取得メイン（起動時・更新ボタン押下時に呼ばれる）
// ==========================================
function getEvents() {
  try {
    // ★アプリ起動時・更新ボタン押下時に、Googleカレンダー側の外部更新・削除を即座に同期
    syncExternalUpdates_();

    const calendar = getTargetCalendar_();
    const now = new Date();
    const year = now.getFullYear();
    
    const startDate = new Date(year, 0, 1);
    const endDate = new Date(year + 1, 11, 31);
    
    const calEvents = calendar.getEvents(startDate, endDate);
    
    const result = calEvents.map(ev => {
      const isAllDay = ev.isAllDayEvent();
      let endStr = ev.getEndTime().toISOString();
      
      if (isAllDay) {
        const adjustedEndDate = ev.getEndTime();
        adjustedEndDate.setDate(adjustedEndDate.getDate() - 1);
        adjustedEndDate.setHours(23, 59, 0, 0);
        endStr = adjustedEndDate.toISOString();
      }

      return {
        id: ev.getId(),
        title: ev.getTitle(),
        start: ev.getStartTime().toISOString(),
        end: endStr,
        location: ev.getLocation(),
        description: ev.getDescription(),
        isAllDay: isAllDay,
        color: ev.getColor()
      };
    });

    const holidays = [];
    for(let y = year; y <= year + 1; y++){
      const holidayMap = getJapaneseHolidays(y);
      for (let key in holidayMap) {
        const parts = key.split('/');
        holidays.push({
          date: `${y}-${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}`,
          title: holidayMap[key]
        });
      }
    }

    const logoBase64 = getLogoBase64_();
    const logs = getLogs_();
    
    return { success: true, events: result, holidays: holidays, logs: logs, profileImage: logoBase64 };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 5. ロゴ画像処理（キャッシュ付きBase64）
// ==========================================
function getLogoBase64_() {
  let logoBase64 = CacheService.getScriptCache().get(LOGO_CACHE_KEY);
  if (logoBase64) return logoBase64;
  const url = getConfigValue_("PROFILE_IMAGE_URL");
  if (!url) return null;
  try {
    const fileId = url.match(/[-\w]{25,}/);
    if (fileId) {
      const file = DriveApp.getFileById(fileId[0]);
      const blob = file.getBlob();
      logoBase64 = "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes());
      CacheService.getScriptCache().put(LOGO_CACHE_KEY, logoBase64, 21600);
      return logoBase64;
    }
  } catch (e) { console.error("ロゴ取得失敗:", e); }
  return null;
}

// ==========================================
// 6. 日本の祝日計算ライブラリ
// ==========================================
function getJapaneseHolidays(year) {
  const holidays = {};
  holidays['1/1'] = '元日'; holidays['2/11'] = '建国記念の日'; holidays['2/23'] = '天皇誕生日';
  holidays['4/29'] = '昭和の日'; holidays['5/3'] = '憲法記念日'; holidays['5/4'] = 'みどりの日';
  holidays['5/5'] = 'こどもの日'; holidays['8/11'] = '山の日'; holidays['11/3'] = '文化の日';
  holidays['11/23'] = '勤労感謝の日';

  const seijin = getNthMondayOfMonth(year, 1, 2); holidays[seijin.m + '/' + seijin.d] = '成人の日';
  const umi = getNthMondayOfMonth(year, 7, 3); holidays[umi.m + '/' + umi.d] = '海の日';
  const keirou = getNthMondayOfMonth(year, 9, 3); holidays[keirou.m + '/' + keirou.d] = '敬老の日';
  const sports = getNthMondayOfMonth(year, 10, 2); holidays[sports.m + '/' + sports.d] = 'スポーツの日';

  const vernalDay = calcVernalEquinoxDay(year); holidays['3/' + vernalDay] = '春分の日';
  const autumnalDay = calcAutumnalEquinoxDay(year); holidays['9/' + autumnalDay] = '秋分の日';

  const allDates = Object.keys(holidays).map(k => {
    const parts = k.split('/'); return { month: parseInt(parts[0]), day: parseInt(parts[1]), key: k };
  });
  allDates.sort((a, b) => (a.month * 100 + a.day) - (b.month * 100 + b.day));

  allDates.forEach(item => {
    const d = new Date(year, item.month - 1, item.day);
    if (d.getDay() === 0) { 
      let subDay = item.day + 1; let subMonth = item.month; let subKey = subMonth + '/' + subDay;
      while (holidays.hasOwnProperty(subKey)) { subDay++; subKey = subMonth + '/' + subDay; }
      holidays[subKey] = '振替休日';
    }
  });

  const finalDates = Object.keys(holidays).map(k => {
    const parts = k.split('/'); return { month: parseInt(parts[0]), day: parseInt(parts[1]) };
  });
  finalDates.sort((a, b) => (a.month * 100 + a.day) - (b.month * 100 + b.day));

  for (let i = 0; i < finalDates.length - 1; i++) {
    const d1 = new Date(year, finalDates[i].month - 1, finalDates[i].day);
    const d2 = new Date(year, finalDates[i + 1].month - 1, finalDates[i + 1].day);
    const diffDays = (d2 - d1) / (1000 * 60 * 60 * 24);
    if (diffDays === 2) {
      const between = new Date(d1); between.setDate(between.getDate() + 1);
      const bKey = (between.getMonth() + 1) + '/' + between.getDate();
      if (between.getDay() !== 0 && !holidays.hasOwnProperty(bKey)) { holidays[bKey] = '国民の休日'; }
    }
  }
  return holidays;
}
function getNthMondayOfMonth(year, month, n) {
  const firstDay = new Date(year, month - 1, 1);
  let day = 1 + ((8 - firstDay.getDay()) % 7);
  if (firstDay.getDay() === 1) day = 1; 
  day += (n - 1) * 7;
  return { m: month, d: day };
}
function calcVernalEquinoxDay(year) {
  if (year >= 1980 && year <= 2099) return Math.floor(20.8431 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  return 20; 
}
function calcAutumnalEquinoxDay(year) {
  if (year >= 1980 && year <= 2099) return Math.floor(23.2488 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  return 23; 
}

// ==========================================
// 7. 予定の操作
// ==========================================
function saveEvent(eventData) {
  try {
    const calendar = getTargetCalendar_();
    let event;
    const start = new Date(eventData.start);
    const end = new Date(eventData.end);
    
    const allDayEnd = new Date(end);
    if (eventData.isAllDay) {
      allDayEnd.setDate(allDayEnd.getDate() + 1);
      allDayEnd.setHours(0, 0, 0, 0); 
    }
    
    let metaTags = "";
    const rMap = { "daily": "毎日", "weekly": "毎週", "monthly": "毎月", "yearly": "毎年" };
    if (eventData.recurrence && eventData.recurrence !== "none") metaTags += `[繰り返し設定: ${rMap[eventData.recurrence]}]\n`;
    const aMap = { "m15": "15分前", "m30": "30分前", "h1": "1時間前", "d1": "1日前" };
    if (eventData.alarm && eventData.alarm !== "none") metaTags += `[通知設定: ${aMap[eventData.alarm]}]\n`;

    const finalDescription = (metaTags + (eventData.description || "")).trim();
    
    if (eventData.id) {
      event = calendar.getEventById(eventData.id);
      if (event) {
        event.setTitle(eventData.title);
        event.setDescription(finalDescription);
        event.setLocation(eventData.location || "");
        if (eventData.isAllDay) event.setAllDayDates(start, allDayEnd);
        else event.setTime(start, end);
      } else { throw new Error("予定が見つかりません。"); }
    } else {
      event = eventData.isAllDay ? 
        calendar.createAllDayEvent(eventData.title, start, allDayEnd, {description: finalDescription, location: eventData.location}) :
        calendar.createEvent(eventData.title, start, end, {description: finalDescription, location: eventData.location});
    }

    if (eventData.flag) {
      if (eventData.flag === "mama") event.setColor("1"); 
      else if (eventData.flag === "papa") event.setColor("10"); 
      else event.setColor("6"); 
    }

    const actionText = eventData.id ? "更新" : "追加";
    
    const actualStartStr = event.getStartTime().toISOString();
    // ★終日予定かどうか(eventData.isAllDay)を引数に追加
    logAction(actionText, eventData.title, actualStartStr, eventData.location, eventData.author, event.getId(), eventData.isAllDay);

    return { success: true, eventId: event.getId() };
  } catch (e) { return { success: false, error: e.toString() }; }
}

function deleteEvent(eventId) {
  try {
    const calendar = getTargetCalendar_();
    const event = calendar.getEventById(eventId);
    if (event) {
      const title = event.getTitle();
      const start = event.getStartTime().toISOString();
      event.deleteEvent();
      // ★終日予定かどうか(event.isAllDayEvent())を引数に追加
      logAction("削除", title, start, "", "歩", eventId, event.isAllDayEvent());
      return { success: true };
    }
    throw new Error("削除対象の予定が見つかりません。");
  } catch (e) { return { success: false, error: e.toString() }; }
}

function cloneEvent(eventId, author) {
  try {
    const calendar = getTargetCalendar_();
    const originalEvent = calendar.getEventById(eventId);
    if (!originalEvent) throw new Error("複製元の予定が見つかりません。");
    
    const newTitle = originalEvent.getTitle() + " (複製)";
    const start = originalEvent.getStartTime();
    const end = originalEvent.getEndTime();
    const desc = originalEvent.getDescription();
    const loc = originalEvent.getLocation();
    
    let newEvent;
    if (originalEvent.isAllDayEvent()) {
      newEvent = calendar.createAllDayEvent(newTitle, start, end, {description: desc, location: loc});
    } else {
      newEvent = calendar.createEvent(newTitle, start, end, {description: desc, location: loc});
    }
    
    const color = originalEvent.getColor();
    if (color) newEvent.setColor(color);
    
    const authorName = author === 'papa' ? '貴弘' : '歩';
    // ★終日予定かどうか(originalEvent.isAllDayEvent())を引数に追加
    logAction("複製", newTitle, start.toISOString(), loc, authorName, newEvent.getId(), originalEvent.isAllDayEvent());
    
    return { success: true, eventId: newEvent.getId() };
  } catch (e) { return { success: false, error: e.toString() }; }
}

// ==========================================
// 8. ログ操作・LINE通知
// ==========================================
function getLogs_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("LOGS");
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    const numRows = Math.min(lastRow - 1, 100);
    const startRow = lastRow - numRows + 1;
    const data = sheet.getRange(startRow, 1, numRows, 7).getValues();
    
    const logs = [];
    for (let i = data.length - 1; i >= 0 && logs.length < 50; i--) {
      const row = data[i];
      // ★お知らせ履歴から「削除」「外部削除」をスキップする設定はすでにここで効いています
      if (row[1] === "削除" || row[1] === "外部削除") continue;
      
      logs.push({
        id: "log_" + (startRow + i),
        timestamp: row[0] ? new Date(row[0]).toISOString() : "",
        action: row[1] || "",
        title: row[2] || "",
        start: row[3] ? new Date(row[3]).toISOString() : "",
        location: row[4] || "",
        author: row[5] || "",
        eventId: row[6] || ""
      });
    }
    return logs;
  } catch (e) { console.error("ログ取得エラー:", e); return []; }
}

// ★第7引数に isAllDay = false を追加
function logAction(action, title, startStr, location, authorId, eventId, isAllDay = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("LOGS");
    if (!sheet) {
      sheet = ss.insertSheet("LOGS");
      sheet.appendRow(["タイムスタンプ", "アクション", "タイトル", "日時", "場所", "登録者", "イベントID"]);
      sheet.setFrozenRows(1);
    }
    
    const authorName = (authorId === 'papa' || authorId === '貴弘') ? '貴弘' : 
                       (authorId === 'mama' || authorId === '歩') ? '歩' : authorId || '不明';
    
    const now = new Date();
    sheet.appendRow([now, action, title, startStr, location, authorName, (eventId || "")]);
    
    // ★LINE通知の条件から "外部削除" と "削除" を除外してミュート化
    if (action === "追加" || action === "更新" || action === "複製" || action === "外部更新") {
      let displayAction = action;
      if (action === "外部更新") displayAction = "追加・更新";

      const dateObj = new Date(startStr);
      // ★終日予定の場合は「終日」、そうでない場合は時刻を生成
      let timeStr = `${String(dateObj.getHours()).padStart(2,'0')}:${String(dateObj.getMinutes()).padStart(2,'0')}`;
      if (isAllDay) {
        timeStr = "終日";
      }
      
      const dateStr = `${dateObj.getMonth()+1}/${dateObj.getDate()} ${timeStr}`;
      
      const msg = `📅 予定が${displayAction}されました\nタイトル: ${title}\n日時: ${dateStr}\n登録: ${authorName}`;
      sendLineNotification(msg, authorName);
    }
  } catch (e) { console.error("ログ記録エラー:", e); }
}

function sendLineNotification(messageText, authorName) {
  if (!messageText) return;

  const channelToken = getConfigValue_("LINE_CHANNEL_ACCESS_TOKEN");
  const isMama = (authorName === "歩" || authorName === "mama");
  const targetIdKey = isMama ? "LINE_USER_ID_PAPA" : "LINE_USER_ID";
  const userId = getConfigValue_(targetIdKey);
  
  if (!channelToken || !userId) return;
  
  const url = 'https://api.line.me/v2/bot/message/push';
  const payload = {
    to: userId,
    messages: [{ type: 'text', text: messageText }]
  };
  
  const options = {
    method: 'post',
    headers: { 'Content-Type': 'application/json', 'Authorization': 'Bearer ' + channelToken },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      console.error("LINE通知送信エラー: HTTP " + responseCode + " - " + response.getContentText());
    }
  } catch (e) {
    console.error("LINE APIリクエスト失敗:", e);
  }
}

// ==========================================
// 10. カレンダー変更の自動検知同期ロジック (更新＆削除検知)
// ==========================================

function syncExternalUpdates_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let lastSyncStr = getConfigValue_("LAST_SYNC_TIME");
    let lastSyncTime = lastSyncStr ? new Date(lastSyncStr).getTime() : new Date().getTime() - (1000 * 60 * 60); 
    
    const calendar = getTargetCalendar_();
    const now = new Date();
    
    const startDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    const endDate = new Date(now.getFullYear(), now.getMonth() + 3, 0);
    
    const events = calendar.getEvents(startDate, endDate);
    
    events.forEach(ev => {
      if (ev.getLastUpdated().getTime() > lastSyncTime) {
        const eventId = ev.getId();
        
        if (!isAlreadyLoggedById_(eventId, lastSyncTime)) {
          const title = ev.getTitle();
          const startStr = ev.getStartTime().toISOString();
          const location = ev.getLocation() || "";
          
          let author = "papa"; 
          if (ev.getColor() === "1") { 
            author = "mama"; 
          }
          
          // ★終日フラグ ev.isAllDayEvent() を追加
          logAction("外部更新", title, startStr, location, author, eventId, ev.isAllDayEvent());
        }
      }
    });

    const currentEventIds = new Set(events.map(ev => ev.getId()));
    
    const sheet = ss.getSheetByName("LOGS");
    if (sheet && sheet.getLastRow() > 1) {
      const numRows = Math.min(sheet.getLastRow() - 1, 150); 
      const startRow = sheet.getLastRow() - numRows + 1;
      const data = sheet.getRange(startRow, 1, numRows, 7).getValues();
      
      const trackedEvents = {};
      for (let i = 0; i < data.length; i++) {
        const id = data[i][6];
        if (id) {
          trackedEvents[id] = {
            action: data[i][1],
            title: data[i][2],
            startStr: data[i][3] ? new Date(data[i][3]).toISOString() : "",
            location: data[i][4],
            author: data[i][5]
          };
        }
      }

      for (const [id, eventInfo] of Object.entries(trackedEvents)) {
        if (eventInfo.action === "削除" || eventInfo.action === "外部削除") continue;
        
        const eventTime = new Date(eventInfo.startStr).getTime();
        if (eventTime >= startDate.getTime() && eventTime <= endDate.getTime()) {
          if (!currentEventIds.has(id)) {
            // ★外部削除時は通知されないため、isAllDayフラグはfalseのままでも影響なし
            logAction("外部削除", eventInfo.title, eventInfo.startStr, eventInfo.location, "papa", id);
          }
        }
      }
    }
    
    updateConfigValue_("LAST_SYNC_TIME", new Date().toISOString());
  } catch (e) {
    console.error("外部更新同期エラー:", e);
  }
}

function checkForCalendarUpdates() {
  syncExternalUpdates_();
}

function isAlreadyLoggedById_(eventId, sinceTime) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("LOGS");
  if (!sheet) return false;
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return false;
  
  const numRows = Math.min(lastRow - 1, 30); 
  const data = sheet.getRange(lastRow - numRows + 1, 1, numRows, 7).getValues();
  
  for (let i = 0; i < data.length; i++) {
    const logTime = new Date(data[i][0]).getTime();
    if (logTime >= sinceTime) {
      if (data[i][6] === eventId && (data[i][1] === "追加" || data[i][1] === "更新")) {
        return true;
      }
    }
  }
  return false;
}

function updateConfigValue_(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName("CONFIG");
  if (!configSheet) return;
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      configSheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  configSheet.appendRow([key, value, "自動同期の最終チェック時間"]);
}

// ==========================================
// ★ これをGASエディタで1回だけ手動実行してください ★
// ==========================================
function setupAutoSyncTrigger() {
  const functionName = "checkForCalendarUpdates";
  const triggers = ScriptApp.getProjectTriggers();
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  ScriptApp.newTrigger(functionName)
    .timeBased()
    .everyMinutes(5)
    .create();
    
  updateConfigValue_("LAST_SYNC_TIME", new Date().toISOString());
  Logger.log("自動同期トリガーを設定しました（5分間隔）。");
  return "自動同期トリガーを設定しました（5分間隔）。";
}

// ==========================================
// ★ LINE通知テスト用関数（権限承認用）★
// ==========================================
function testLineNotification() {
  const channelToken = getConfigValue_("LINE_CHANNEL_ACCESS_TOKEN");
  const mamaId = getConfigValue_("LINE_USER_ID");
  const papaId = getConfigValue_("LINE_USER_ID_PAPA");
  
  Logger.log("【設定チェック】");
  Logger.log("チャネルトークン: " + (channelToken ? "入力あり" : "未入力"));
  Logger.log("LINE_USER_ID (歩): " + (mamaId ? "入力あり" : "未入力"));
  Logger.log("LINE_USER_ID_PAPA (貴弘): " + (papaId ? "入力あり" : "未入力"));

  const msg = "【テスト通知】AYUMY SyncCalからのテストです💜\nこのメッセージが届いていれば設定は完璧です！";
  
  Logger.log("--- 貴弘(PAPA)宛へのテスト送信開始 ---");
  sendLineNotification(msg, "歩"); // 歩として送信＝貴弘(PAPA)に届くはず
  
  Logger.log("--- 歩(MAMA)宛へのテスト送信開始 ---");
  sendLineNotification(msg, "貴弘"); // 貴弘として送信＝歩(MAMA)に届くはず
  
  Logger.log("テスト通知の処理が完了しました。上記にHTTP 400エラーが出ている場合、指定したIDのユーザーがLINE Botを友だち追加していないか、IDが間違っています。");
  return "テスト実行完了";
}
