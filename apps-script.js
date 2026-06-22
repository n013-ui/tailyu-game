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
 *
 * ★ 更新步驟（已部署過的話）：
 *    點擊「部署」→「管理部署」→ 編輯（鉛筆圖示）→ 版本選「新版本」→ 部署
 *    部署完成後，在瀏覽器開啟：
 *    https://script.google.com/macros/s/你的ID/exec?action=setup
 *    等看到 {"success":true} 即完成初始化
 */

const SHEET_ID = '10anvoOIaduohZR0W1vYDQbZRiYTo1FbXHJ9KSz_FR_I';

// ══════════════════════════════════════════════════
//  試算表結構（執行 ?action=setup 自動建立 / 升級）
//  題目設定欄位：關卡號│類型│名稱│Emoji│題目1│題目2│題目3│題目4│分數│顏色
// ══════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ['題目設定','機會卡','命運卡','遊戲進度'].forEach(name => {
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });

  // ── 題目設定 ──
  const q = ss.getSheetByName('題目設定');
  // 舊格式（欄數不足）或格數 < 36（含標題列）就清除重建
  const curName2 = q.getLastRow() > 2 ? String(q.getRange(3,3,1,1).getValue()) : '';
  const oldFormat = q.getLastRow() === 0 || q.getLastColumn() < 8 || q.getLastRow() < 37 || curName2 !== '俚語題';
  if (oldFormat) {
    q.clearContents();
    q.appendRow(['關卡號','類型','名稱','Emoji','題目1','題目2','題目3','題目4','分數','顏色']);
    [
      [1,'start','起點','🚀',
       '遊戲開始！','','','',
       0,'#276749'],

      [2,'task','俚語題','💬',
       '說出1個和「動物」有關的台語俚語！','說出1個和「植物」有關的台語俚語！','說出1個和「人」有關的台語俚語！','說出1個和「態度」有關的台語諺語！',
       10,'#2b6cb0'],

      [3,'song','兒歌題','🎵',
       '用台語唱一首兒歌','用台語唱一首兒歌(不可重複)','用台語唱一首兒歌(不可重複)','用台語唱一首兒歌(不可重複)',
       15,'#c05621'],

      [4,'chance','機　會','⭐','抽一張機會卡！','','','', 0,'#b7791f'],

      [5,'riddle','詞彙題','🔍',
       '\n「一個人做事情笨手笨腳，或不擅長做某件事時」\n我們可以用什麼詞來形容 \n請在黑板寫出來台語正字，並用台語唸出來',
       '\n「開玩笑」的台語是什麼？\n有兩種說法，全對才給分…  \n請在黑板寫出來台語正字，並用台語唸出來',
       '\n「醫院」的台語是什麼？\n請在黑板寫出來台語正字，並用台語唸出來',
       '\n「什麼」的台語常被誤用為  蝦米\n請在黑板寫出來台語正字，並用台語唸出來',
       10,'#6b46c1'],

      [6,'task','繞口令','👄',
       '一隻猴仔帶一寡猴仔去溝仔，\n一隻猴仔跋落溝仔，\n一隻猴仔轉去提鉤仔來鉤猴仔。',
       '我共你講喔！你莫共別人講！\n你若欲共別人講，我就無愛共你講！\n你若共別人講！\n你毋通講是我共你講！\n你莫共別人講！\n無我袂閣再共你講！',
       '\n樹頂一隻猴，樹跤一隻狗，\n猴看著狗，狗看著猴，\n猴走狗嘛走，\n毋知是狗驚猴，抑是猴驚狗',
       '囡仔囡仔弄破鼓，提褲來補鼓，\n鼓破褲補，褲破布補，\n到底是布補鼓抑是褲補鼓。',
       15,'#2b6cb0'],

      [7,'fortune','命　運','🎴','抽一張命運卡！','','','', 0,'#c53030'],

      [8,'task','拼音題','📚',
       '寫出五個「p」開頭的台語字(詞)','寫出五個「m」開頭的台語字(詞)','寫出五個「th」開頭的台語字(詞)','寫出五個「s」開頭的台語字(詞)',
       10,'#2b6cb0'],

      [9,'task','情境','🎭',
       '去菜市場想要買魚，\n要怎麼用台語跟老闆詢問價格？',
       '\n去餐廳吃飯…\n看完菜單，要怎麼點菜 \n(請用台語點三道菜)',
       '走在路上看到正妹，\n要怎麼跟她跟她要電話號碼？\n(請用台語回答)',
       '去看醫生的時候，醫生詢問你怎麼了？\n請你用台語說出自己目前感冒的症狀',
       10,'#2b6cb0'],

      [10,'task','數數','🔢',
       '用台語從 1 數到 20！\n一、二、三…念得流暢才算過關！',
       '用台語說出今天的日期！\n年、月、日都要說出來！',
       '小明從今天開始每天存錢，第一天存1元，\n第二天2元…存到第十天，一共有多少錢？\n(使用台語回答)',
       '今仔日是拜五，昨昏是拜四 ，昨日是拜幾？',
       10,'#2b6cb0'],

      [11,'chance','機　會','⭐','抽一張機會卡！','','','', 0,'#b7791f'],

      [12,'task','親情','💝',
       '舅媽的台語','伯母的台語','姨丈的台語','姑丈的台語',
       10,'#2b6cb0'],

      [13,'song','支援前線','🎶',
       '一個同學負責回答\n由同組同學拿出五項東西\n需正確回答五項東西的台語才能得分',
       '一個同學負責回答\n由同組同學拿出五項東西\n需正確回答五項東西的台語才能得分',
       '一個同學負責回答\n由同組同學拿出五項東西\n需正確回答五項東西的台語才能得分',
       '一個同學負責回答\n由同組同學拿出五項東西\n需正確回答五項東西的台語才能得分',
       15,'#c05621'],

      [14,'fortune','命　運','🎴','抽一張命運卡！','','','', 0,'#c53030'],

      [15,'task','台語方向','🧭',
       '每組選三個同學出來，\n閉眼聽完老師的口令轉動身體，\n口令分為「越正爿」、「越左爿」、「向後壁」、「向頭前」\n最後三人面向相同的方向即得分',
       '每組選三個同學出來，\n閉眼聽完老師的口令轉動身體，\n口令分為「越正爿」、「越左爿」、「向後壁」、「向頭前」\n最後三人面向相同的方向即得分',
       '每組選三個同學出來，\n閉眼聽完老師的口令轉動身體，\n口令分為「越正爿」、「越左爿」、「向後壁」、「向頭前」\n最後三人面向相同的方向即得分',
       '每組選三個同學出來，\n閉眼聽完老師的口令轉動身體，\n口令分為「越正爿」、「越左爿」、「向後壁」、「向頭前」\n最後三人面向相同的方向即得分',
       10,'#2b6cb0'],

      [16,'task','食物','🍜',
       '用台語說出 3 種台灣夜市食物！','用台語說出 3 種台灣傳統點心！','用台語說出 3 種過年會吃的食物！','用台語說出 3 種飲料',
       10,'#2b6cb0'],

      [17,'chance','機　會','⭐','抽一張機會卡！','','','', 0,'#b7791f'],

      [18,'task','動物','🐾',
       '用台語說出五種可以在海底游的生物','用台語說出五種會養在農場的動物','用台語說出五種會長在樹上的水果','用台語說出五種拜拜時供桌上會出現的食物',
       10,'#2b6cb0'],

      [19,'fortune','命　運','🎴','抽一張命運卡！','','','', 0,'#c53030'],

      [20,'task','問候','🤝',
       '用台語說出三句跟長輩可以用的問候語',
       '用台語說出過年會說的吉祥話，至少三句',
       '早上上學看到同學、老師可以怎麼問候？\n使用台語，至少三句',
       '回到家，看到爸媽可以怎麼跟家人問候？\n(使用台語，至少二句)',
       10,'#2b6cb0'],

      [21,'chance','機　會','⭐','抽一張機會卡！','','','', 0,'#b7791f'],

      [22,'task','台語身體','🧍',
       '用台語說出 5 個會出現在頭部的器官或部位！',
       '用台語從頭到腳說出 5 個身體部位！\n(順序必須要從由上而下)',
       '「囡仔人有耳無喙」，是什麼意思？\n(本題可以用中文回答)',
       '指指 kí-tsáinn  是五支手指頭的哪一支？',
       10,'#2b6cb0'],

      [23,'song','植物','🌱',
       '用台語說出三種菜市場會賣的菜(要有葉子的)','用台語說出會長在土裡的植物五種','用台語說出三種菜市場會賣的瓜類','用台語說出三種花的名字',
       15,'#c05621'],

      [24,'fortune','命　運','🎴','抽一張命運卡！','','','', 0,'#c53030'],

      [25,'task','選擇題','❓',
       '這馬西瓜當著時，是指西瓜生產的情況按怎？\n(A)大出 (B)欲無矣',
       '伊足gau5「囥歲」是指？伊看起來真…？\n(A)少年 (B)臭老',
       '這个人真gau5「張」是講伊真…\n(A) 歹性地   (B)真骨力',
       '這个人真gau5「趖」是講伊做代誌真？\n(A) 貧惰  (B)慢',
       10,'#2b6cb0'],

      [26,'task','動物','🦒',
       '物件食食咧愛收予好，才袂生______？\n_________要填入什麼動物？(台語)',
       '長頸鹿的台語怎麼說？','企鵝的台語怎麼說？',
       '卡皮巴拉就是水肫君，他的台語是？\n(A) 水豬  (B)水鼠',
       10,'#2b6cb0'],

      [27,'song','習俗','🏮',
       '用台語說出二種我們會在端午節做的事','用台語說出二種我們會在中秋節做的事','用台語說出二種我們會在元宵節做的事','用台語說出二種我們會在清明節做的事',
       15,'#c05621'],

      [28,'task','環境','🏠',
       '用台語說出房間裡會有的東西，五種！','用台語說出廚房裡會有的東西，五種！','用台語說出廁所裡會有的東西，五種！','用台語說出客廳裡會有的東西，五種！',
       10,'#2b6cb0'],

      [29,'task','台語家庭','👨‍👩‍👧',
       '用台語說出小丸子家中所有成員的稱謂\n爸爸、媽媽、阿公、阿媽、姊姊',
       '用四句台語形容你的家庭',
       '說出媽媽通常會有的三種特質\n(用台語回答)',
       '說出爸爸通常會有的三種特質\n(用台語回答)',
       10,'#2b6cb0'],

      [30,'task','語錄','💭',
       '用台語說出導師常會說的三句話',
       '用台語說出媽媽常會說的三句話',
       '用台語說出爸爸常會說的三句話',
       '用台語說出一句經典的廣告詞，或是電視劇的台詞\n【要說明是什麼廣告或什麼電視劇】',
       10,'#2b6cb0'],

      [31,'task','學校題','🏫',
       '用台語說出 3 種學校課程名稱！','學校的校鴨品種是？(以台語回答)','用台語說出教室裡的 5 種物品！','用台語說出「南崁國中」四個字',
       10,'#2b6cb0'],

      [32,'song','課本','📖',
       '拜訪親友時帶的禮物叫做？\n(請用台語回答)',
       '勸人做事情要小心、仔細、謹慎要用什麼詞？\n(請用台語回答)',
       '他已經冷到直「發抖」，\n我們台語可以怎麼說？',
       '他是一個害羞內向的人，\n我們台語可以怎麼說？',
       15,'#c05621'],

      [33,'task','準備','🛒',
       '要去烤肉時，我們會買的五樣東西 \n(用台語回答)',
       '要去大便時，我們會準備的三樣東西 \n(用台語回答)',
       '防颱準備時，我們會準備的三樣東西 \n(用台語回答)',
       '要去考試時，我們會準備的五樣東西 \n(用台語回答)',
       10,'#2b6cb0'],

      [34,'task','接龍','🔗',
       '使用「毋甘」開頭，再接二個詞\n(使用台語作答)',
       '使用「烘肉」開頭，再接二個詞\n(使用台語作答)',
       '使用「頭路」開頭，再接二個詞\n(使用台語作答)',
       '使用「後擺」開頭，再接二個詞\n(使用台語作答)',
       10,'#2b6cb0'],

      [35,'chance','機　會','⭐','抽一張機會卡！','','','', 0,'#b7791f'],

      [36,'end','終　點','🏆',
       '恭喜到達終點！','','','', 20,'#b7791f'],
    ].forEach(row => q.appendRow(row));
  }

  // ── 機會卡 ──
  const ch = ss.getSheetByName('機會卡');
  if (ch.getLastRow() === 0) {
    ch.appendRow(['編號','Emoji','卡片內容','效果類型','效果數值','效果說明']);
    [
      [1,'🍀','幸運之神眷顧！\n前進 2 格！','move',2,'前進 2 格'],
      [2,'⭐','台語說得真棒！\n獲得 15 分！','score',15,'獲得 15 分'],
      [3,'🎲','再來一次！\n可以再擲一次骰子！','extra_roll',1,'再擲一次骰子'],
      [4,'🌟','老師嘉獎！\n全部組別各獲得 5 分！','all_score',5,'全組各得 5 分'],
      [5,'🚀','發現捷徑！\n直接傳送到第 8 格！','goto',8,'傳送到第 8 格'],
      [6,'🤝','同學互助！\n選任一組也獲得 10 分！','give_score',10,'贈送 10 分'],
    ].forEach(row => ch.appendRow(row));
  }

  // ── 命運卡 ──
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

  // ── 遊戲進度 ──
  const pr = ss.getSheetByName('遊戲進度');
  if (pr.getLastRow() === 0) {
    pr.appendRow(['班級','儲存名稱','儲存時間','隊伍數','遊戲狀態JSON']);
  }

  return ContentService.createTextOutput(JSON.stringify({success:true, message:'試算表已設定完成（新版 4 題格式）'}))
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
      default: result = {ok:true, message:'台語大富翁 API 運作中（4 題版）'};
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
//  QUESTIONS  （新格式：10 欄，題目1~4 各一欄）
// ══════════════════════════════════════════════════
function getQuestions() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const result = {};

  const qSheet = ss.getSheetByName('題目設定');
  if (qSheet && qSheet.getLastRow() > 1) {
    const rows = qSheet.getRange(2, 1, qSheet.getLastRow()-1, 10).getValues();
    result.questions = rows
      .filter(r => r[0])
      .map(r => ({
        n:         Number(r[0]),
        type:      String(r[1]),
        name:      String(r[2]),
        emoji:     String(r[3]),
        questions: [r[4], r[5], r[6], r[7]]
                     .map(v => String(v).trim())
                     .filter(Boolean),   // 空格自動跳過
        pts:       Number(r[8]) || 10,
        color:     String(r[9]) || '#2b6cb0',
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

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(classId) && String(data[i][1]) === String(name)) {
      sheet.getRange(i+1, 3, 1, 3).setValues([[now, countTeams(stateJson), stateJson]]);
      return {success: true, updated: true};
    }
  }
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
