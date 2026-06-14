/**
 * 台語大富翁 - Google Apps Script 後端
 *
 * 請照以下步驟設定：
 * 1. 建立一個新的 Google 試算表
 * 2. 開啟「擴充功能」→「Apps Script」
 * 3. 將此檔案全部內容貼入編輯器
 * 4. 修改下方 SHEET_ID 為你的試算表 ID（網址中的長串代碼）
 * 5. 點擊「部署」→「新部署」→「類型：網頁應用程式」
 *    - 誰可以執行：所有人
 *    - 存取權：所有人（匿名）
 * 6. 複製部署網址，貼到遊戲的「設定 → Apps Script 網址」
 */

const SHEET_ID = '10anvoOIaduohZR0W1vYDQbZRiYTo1FbXHJ9KSz_FR_I';

// ══════════════════════════════════════════════════
//  試算表結構（第一次使用時自動建立）
// ══════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const needed = ['題目設定', '機會卡', '命運卡', '遊戲進度'];
  needed.forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  // 題目設定 headers
  const q = ss.getSheetByName('題目設定');
  if (q.getLastRow() === 0) {
    q.appendRow(['關卡號', '類型', '名稱', 'Emoji', '任務說明', '分數', '顏色']);
    // 預設20關資料
    const defaults = [
      [1,'start','起點','🚀','遊戲開始！大家一起說台語！',0,'#276749'],
      [2,'task','台語俚語','💬','說出3個台語俚語\n例如：食果子拜樹頭、囝仔人有耳無喙',10,'#2b6cb0'],
      [3,'song','台語歌謠','🎵','用台語唱一首歌謠，至少唱4句',15,'#c05621'],
      [4,'chance','機會','⭐','抽一張機會卡！',0,'#b7791f'],
      [5,'riddle','台語猜謎','🔍','謎題：「一粒珠，圓閣圓」用台語說出謎底（月亮）',10,'#6b46c1'],
      [6,'task','台語繞口令','👄','說台語繞口令三遍：「是獅是虎是獅虎」',15,'#2b6cb0'],
      [7,'fortune','命運','🎴','抽一張命運卡！',0,'#c53030'],
      [8,'task','台語成語','📚','說出一個台語成語並解釋意思',10,'#2b6cb0'],
      [9,'rest','休息一下','😴','休息一回合！下一次輪到你時跳過。',0,'#4a5568'],
      [10,'task','台語數數','🔢','用台語從1數到20，念得流暢才過關！',10,'#2b6cb0'],
      [11,'chance','機會','⭐','抽一張機會卡！',0,'#b7791f'],
      [12,'task','台語自我介紹','👋','用台語做自我介紹（名字、年紀、學校、喜好）',10,'#2b6cb0'],
      [13,'song','台語童謠','🎶','表演台語童謠，至少唱完一個完整段落',15,'#c05621'],
      [14,'fortune','命運','🎴','抽一張命運卡！',0,'#c53030'],
      [15,'task','台語方向','🧭','用台語說出東西南北、左右前後',10,'#2b6cb0'],
      [16,'task','台語食物','🍜','用台語說出5種傳統台灣食物名稱',10,'#2b6cb0'],
      [17,'chance','機會','⭐','抽一張機會卡！',0,'#b7791f'],
      [18,'task','台語動物','🐾','用台語說出5種動物名稱並模仿叫聲',10,'#2b6cb0'],
      [19,'fortune','命運','🎴','抽一張命運卡！',0,'#c53030'],
      [20,'end','終點','🏆','恭喜到達終點！你是台語高手！',20,'#b7791f'],
    ];
    defaults.forEach(row => q.appendRow(row));
  }

  // 機會卡 headers
  const ch = ss.getSheetByName('機會卡');
  if (ch.getLastRow() === 0) {
    ch.appendRow(['編號','Emoji','卡片內容','效果類型','效果數值','效果說明']);
    [
      [1,'🍀','幸運之神眷顧！\n前進 2 格！','move',2,'前進 2 格'],
      [2,'⭐','台語說得真棒！\n獲得 15 分！','score',15,'獲得 15 分'],
      [3,'🎲','再來一次！\n可以再擲一次骰子！','extra_roll',1,'再擲一次骰子'],
      [4,'🌟','老師嘉獎！\n全部組別各獲得 5 分！','all_score',5,'全組各得 5 分'],
      [5,'🚀','發現捷徑！\n直接前進到第 15 格！','goto',15,'傳送到第 15 格'],
      [6,'🤝','同學互助！\n選任一組也獲得 10 分！','give_score',10,'贈送 10 分'],
    ].forEach(row => ch.appendRow(row));
  }

  // 命運卡 headers
  const fo = ss.getSheetByName('命運卡');
  if (fo.getLastRow() === 0) {
    fo.appendRow(['編號','Emoji','卡片內容','效果類型','效果數值','效果說明']);
    [
      [1,'😰','走錯路了！\n後退 2 格！','move',-2,'後退 2 格'],
      [2,'😅','台語說錯了！\n失去 5 分！','score',-5,'失去 5 分'],
      [3,'⏸','需要多加練習！\n跳過下一回合！','skip',1,'跳過一回合'],
      [4,'🔀','命運大逆轉！\n與目前第一名交換位置！','swap_first',0,'與第一名換位'],
      [5,'🎉','意外的好事！\n前進 3 格！','move',3,'前進 3 格'],
      [6,'🎤','特別任務！\n用台語唱一句歌詞成功得 20 分！','bonus_task',20,'特別表演任務'],
    ].forEach(row => fo.appendRow(row));
  }

  // 遊戲進度 headers
  const pr = ss.getSheetByName('遊戲進度');
  if (pr.getLastRow() === 0) {
    pr.appendRow(['班級','儲存名稱','儲存時間','隊伍數','遊戲狀態JSON']);
  }

  return ContentService.createTextOutput(JSON.stringify({success:true, message:'試算表已設定完成'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════
//  HTTP 進入點
// ══════════════════════════════════════════════════
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';
  let result;
  try {
    switch(action) {
      case 'getQuestions': result = getQuestions(); break;
      case 'loadProgress': result = loadProgress(e.parameter.classId, e.parameter.name); break;
      case 'listSaves':    result = listSaves(e.parameter.classId); break;
      case 'setup':        return setupSheets();
      default: result = {ok:true, message:'台語大富翁 API 運作中'};
    }
  } catch(err) {
    result = {error: err.message};
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'saveProgress') {
      result = saveProgress(data.classId, data.name, data.state);
    } else {
      result = {error: 'Unknown action'};
    }
  } catch(err) {
    result = {error: err.message};
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════
//  QUESTIONS
// ══════════════════════════════════════════════════
function getQuestions() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const result = {};

  const qSheet = ss.getSheetByName('題目設定');
  if (qSheet && qSheet.getLastRow() > 1) {
    const rows = qSheet.getRange(2, 1, qSheet.getLastRow()-1, 7).getValues();
    result.questions = rows
      .filter(r => r[0])
      .map(r => ({
        n:     Number(r[0]),
        type:  String(r[1]),
        name:  String(r[2]),
        emoji: String(r[3]),
        desc:  String(r[4]),
        pts:   Number(r[5]) || 10,
        color: String(r[6]) || '#2b6cb0',
      }));
  }

  const chSheet = ss.getSheetByName('機會卡');
  if (chSheet && chSheet.getLastRow() > 1) {
    const rows = chSheet.getRange(2, 1, chSheet.getLastRow()-1, 6).getValues();
    result.chanceCards = rows.filter(r=>r[0]).map(r=>({
      emoji: String(r[1]), text: String(r[2]), effect: String(r[3]),
      val: Number(r[4]), label: String(r[5]),
    }));
  }

  const foSheet = ss.getSheetByName('命運卡');
  if (foSheet && foSheet.getLastRow() > 1) {
    const rows = foSheet.getRange(2, 1, foSheet.getLastRow()-1, 6).getValues();
    result.fortuneCards = rows.filter(r=>r[0]).map(r=>({
      emoji: String(r[1]), text: String(r[2]), effect: String(r[3]),
      val: Number(r[4]), label: String(r[5]),
    }));
  }

  return result;
}

// ══════════════════════════════════════════════════
//  SAVE / LOAD PROGRESS
// ══════════════════════════════════════════════════
function saveProgress(classId, name, stateJson) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('遊戲進度');
  if (!sheet) return {success: false, error: '找不到「遊戲進度」工作表'};

  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm:ss');
  const data = sheet.getDataRange().getValues();

  // Check if already exists (same class + name), update in place
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(classId) && String(data[i][1]) === String(name)) {
      sheet.getRange(i+1, 3, 1, 3).setValues([[now, countTeams(stateJson), stateJson]]);
      return {success: true, updated: true};
    }
  }

  // New row
  sheet.appendRow([classId, name, now, countTeams(stateJson), stateJson]);
  return {success: true, inserted: true};
}

function countTeams(stateJson) {
  try { return JSON.parse(stateJson).teams.length; } catch(e) { return 0; }
}

function listSaves(classId) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('遊戲進度');
  if (!sheet || sheet.getLastRow() < 2) return {saves:[]};
  const data = sheet.getRange(2,1,sheet.getLastRow()-1,4).getValues();
  const saves = data
    .filter(r => !classId || String(r[0]) === String(classId))
    .map(r => ({classId:r[0], name:r[1], time:r[2], teams:r[3]}));
  return {saves};
}

function loadProgress(classId, name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('遊戲進度');
  if (!sheet || sheet.getLastRow() < 2) return {error:'找不到資料'};
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(classId) && String(data[i][1]) === String(name)) {
      try {
        return {success: true, state: JSON.parse(data[i][4])};
      } catch(e) {
        return {error:'資料格式錯誤'};
      }
    }
  }
  return {error:'找不到指定的儲存進度'};
}
