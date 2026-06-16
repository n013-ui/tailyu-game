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
  // 若是舊格式（7欄以下）或空白，清除重建
  const oldFormat = q.getLastRow() === 0 || q.getLastColumn() < 8;
  if (oldFormat) {
    q.clearContents();
    q.appendRow(['關卡號','類型','名稱','Emoji','題目1','題目2','題目3','題目4','分數','顏色']);
    [
      // ─ 格式：[n, type, name, emoji, q1, q2, q3, q4, pts, color] ─
      // 機會/命運/休息/起終點：q2~q4 留空即可

      [1,'start','起點','🚀',
       '遊戲開始！\n大家一起說：\n「逐家好！咱來學台語！」\n（Tak-ke hó! Lán lâi o̍h Tâi-gí!）',
       '','','', 0,'#276749'],

      [2,'task','台語俚語','💬',
       '說出 3 個台語俚語並解釋意思！\n‧ 食果子，拜樹頭\n‧ 囝仔人有耳無喙\n‧ 放牛食草',
       '說出 3 個和「做人」有關的台語俚語！\n‧ 好心有好報\n‧ 人情留一線，日後好相見\n‧ 毋通害人害己',
       '說出 3 個和動物有關的台語諺語！\n‧ 虎死留皮，人死留名\n‧ 狗咬狗骨頭\n‧ 雞母帶雞仔\n說說意思！',
       '說出 3 個和「天氣」有關的台語俚語！\n‧ 春天後母面\n‧ 早霞暗，晚霞晴\n‧ 六月無雲是大旱\n說說意思！',
       10,'#2b6cb0'],

      [3,'song','台語歌謠','🎵',
       '用台語唱「一隻鳥仔哮救救」！\n至少唱完一個完整段落，大聲唱！',
       '用台語唱「天黑黑」！\n（天黑黑，欲落雨，阿公仔舉鋤頭欲掘芋…）\n至少 4 句！',
       '用台語唱「望春風」！\n（獨夜無伴守燈下，清風對面吹…）\n至少 4 句！',
       '用台語唱「農村曲」！\n（透早就出門，天色漸漸光…）\n至少 4 句！',
       15,'#c05621'],

      [4,'chance','機　會','⭐','抽一張機會卡！','','','', 0,'#b7791f'],

      [5,'riddle','台語猜謎','🔍',
       '謎題：\n「一粒珠，圓閣圓，大人囡仔攏愛看。」\n用台語說出謎底！（答案：月娘）',
       '謎題：\n「紅門白壁鑽出去，搖頭擺尾弄撥撥。」\n用台語說出謎底！（答案：舌頭）',
       '謎題：\n「有人行，無人搬，行到厝裡才知寒。」\n用台語說出謎底！（答案：風）',
       '謎題：\n「一條龍，無頭無尾，暝日攏在厝裡睏。」\n用台語說出謎底！（答案：水管）',
       10,'#6b46c1'],

      [6,'task','台語繞口令','👄',
       '說台語繞口令，連說 3 遍！\n「是獅是虎是獅虎，\n 是虎是獅是虎獅」',
       '說台語繞口令，連說 3 遍！\n「大廟後有大土地，\n 小廟後有小土地」',
       '說台語繞口令，連說 3 遍！\n「先生先生，先生說，\n 先生說先生念書先生」',
       '說台語繞口令，連說 3 遍！\n「買冰買鳥冰，買鳥買冰冰，\n 毋是鳥冰是冰鳥」',
       15,'#2b6cb0'],

      [7,'fortune','命　運','🎴','抽一張命運卡！','','','', 0,'#c53030'],

      [8,'task','台語成語','📚',
       '說出一個台語成語並解釋意思！\n例如：一人一家代、毋通糊塗',
       '用台語說一句關於「勤勞」的諺語！\n例如：天公疼勤人、早起的鳥仔有蟲吃',
       '用台語說一句關於「友情」的諺語！\n例如：有朋友遠方來、朋友是第二個家',
       '用台語說一句關於「家庭」的諺語！\n例如：家和萬事興、父母恩深似海\n說說意思！',
       10,'#2b6cb0'],

      [9,'rest','休息一下','😴',
       '休息一回合！下一次輪到你時跳過。\n大家用台語說：「好啦！歇睏一下！」',
       '','','', 0,'#4a5568'],

      [10,'task','台語數數','🔢',
       '用台語從 1 數到 20！\n一、二、三…念得流暢才算過關！',
       '用台語說出今天的日期！\n年、月、日都要說出來！',
       '用台語從 20 倒數到 1！\n二十、十九、十八…一！念得流暢才過關！',
       '用台語說出目前時刻、星期幾、今天天氣如何！\n全部說對才過關！',
       10,'#2b6cb0'],

      [11,'chance','機　會','⭐','抽一張機會卡！','','','', 0,'#b7791f'],

      [12,'task','台語自介','👋',
       '用台語做自我介紹！\n‧ 名字 ‧ 年紀 ‧ 學校班級 ‧ 一個喜好',
       '用台語介紹你的家人！\n至少介紹三位（爸爸、媽媽、兄弟姊妹…）',
       '用台語說說你的興趣和夢想！\n說一件最喜歡的事，以及長大想做什麼',
       '用台語說出最喜歡的科目和最不喜歡的科目！\n各說原因，至少 3 句台語！',
       10,'#2b6cb0'],

      [13,'song','台語童謠','🎶',
       '表演台語童謠「丟丟銅仔」！\n（丟丟銅仔，伊都丟，阿末仔伊都丟串的路…）\n至少唱一段！',
       '表演台語童謠「點仔膠」！\n（點仔膠，黏著腳，叫阿爸，買豬腳…）\n唱得愈完整愈好！',
       '表演台語童謠「白鷺鷥」！\n（白鷺鷥，車畚箕，車到溪仔邊…）\n至少唱一段！',
       '表演台語童謠「火金姑」！\n（火金姑，來食茶，茶燒燒，配芋糕…）\n至少唱一段！',
       15,'#c05621'],

      [14,'fortune','命　運','🎴','抽一張命運卡！','','','', 0,'#c53030'],

      [15,'task','台語方向','🧭',
       '用台語說出東西南北、左右前後！\n全部說正確才算！',
       '老師隨機指方向，你用台語喊出來！\n東西南北左右前後，全部答對才過關！',
       '用台語說出「上下內外遠近」六個方位詞！\n說出來並用手比劃！',
       '用台語說明從教室到校門口的路線！\n至少說出 3 個方向詞！',
       10,'#2b6cb0'],

      [16,'task','台語食物','🍜',
       '用台語說出 5 種傳統台灣食物名稱！\n（牛肉麵、滷肉飯、珍珠奶茶、鹽酥雞…）',
       '用台語說出 5 種台灣夜市食物！\n並說說你最愛哪一種，用台語說理由！',
       '用台語說出食物的五種味道！\n甜、酸、苦、辣、鹹\n各舉一種食物例子！',
       '用台語說出 5 種台灣傳統點心！\n（湯圓、麻糬、紅龜粿、草仔粿…）\n說說在什麼時候吃！',
       10,'#2b6cb0'],

      [17,'chance','機　會','⭐','抽一張機會卡！','','','', 0,'#b7791f'],

      [18,'task','台語動物','🐾',
       '用台語說出 5 種動物名稱並模仿牠們的叫聲！\n（狗→汪汪、貓→喵喵、雞→咕咕…）',
       '用台語說出 5 種農場動物！\n並說說牠們住哪裡、吃什麼！',
       '用台語說出 5 種海洋生物！\n（魚、蝦、蟳、花枝、鯊魚）\n並模仿一種牠的游泳方式！',
       '用台語說出 5 種台灣本土動物！\n（台灣黑熊、山豬、梅花鹿、台灣獼猴…）\n說說牠們住在哪裡！',
       10,'#2b6cb0'],

      [19,'fortune','命　運','🎴','抽一張命運卡！','','','', 0,'#c53030'],

      [20,'task','台語問候','🤝',
       '用台語說出早中晚的問候語！\n‧ 早起好！（早安）\n‧ 食飽未？（午安問候）\n‧ 恁好！（你好）',
       '用台語演練買東西的對話！\n「這個偌濟錢？」\n「有較便宜的無？」\n「多謝！」',
       '用台語問路！\n「請問廁所佇佗位？」\n「請問圖書館按怎去？」\n說說看！',
       '用台語說出節慶問候語！\n‧ 新年恭喜！\n‧ 生日快樂！\n‧ 恭喜發財！',
       10,'#2b6cb0'],

      [21,'chance','機　會','⭐','抽一張機會卡！','','','', 0,'#b7791f'],

      [22,'task','台語身體','🫀',
       '用台語說出 7 個頭部器官！\n頭、目睭、耳、鼻、喙、頭毛、面',
       '用台語從頭到腳說出 10 個身體部位！\n說得流暢才算過關！',
       '用台語說出 5 種情緒！\n快樂、悲傷、生氣、驚嚇、厭煩\n各造一個台語例句！',
       '老師指身體部位，你立刻用台語說出！\n連答 8 個才過關！',
       10,'#2b6cb0'],

      [23,'song','台語歌謠','🎤',
       '用台語唱「西北雨」！\n（西北雨，直直落，鯽仔魚，欲娶某…）\n至少唱 4 句！',
       '用台語唱「六月茉莉」！\n（六月茉莉真正芳，白白的花蕊…）\n至少唱 4 句！',
       '用台語唱「月光光」！\n（月光光，照地堂，年三十，摘檳榔…）\n至少唱 4 句！',
       '用台語唱任何一首台語歌！\n要唱完至少 4 句，大家一起評分！',
       15,'#c05621'],

      [24,'fortune','命　運','🎴','抽一張命運卡！','','','', 0,'#c53030'],

      [25,'end','終　點','🏆',
       '恭喜到達終點！\n請用台語說：\n「多謝！咱攏是台語高手！」\n大家一起鼓掌！',
       '','','', 20,'#b7791f'],
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
