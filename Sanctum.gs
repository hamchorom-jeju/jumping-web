/**
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * Sanctum Module (이장의 집 - 기후/지식/공지/퀘스트 통제 백엔드)
 */



/**
 * ✍️ 지혜의 보물고 (집필실) 관련 로직
 */
function getWisdomTips() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("지혜의_보물고") || ss.insertSheet("지혜의_보물고");
    if (sheet.getLastRow() < 2) return []; // 1행은 헤더이므로 데이터는 최소 2행부터 존재함
    
    var data = sheet.getDataRange().getDisplayValues();
    
    var pKeyRes = getPollinationsApiKey();
    var pKey = pKeyRes.success ? pKeyRes.key : "";
    
    var tips = [];
    for (var i = 1; i < data.length; i++) {
      var rawImage = data[i].length > 8 ? data[i][8] : "";
      var finalImage = appendPollinationsApiKey(rawImage, pKey);
      
      tips.push({
        id: i + 1, // 1-indexed 스프레드시트 실제 행 주소로 매칭 (데이터 첫행은 2행이므로 i=1 일때 id=2)
        date: data[i][0],
        title: data[i][1],
        content: data[i][2],
        category: data[i][3],
        author: data[i][4],
        image: finalImage,
        isDefault: data[i].length > 9 ? data[i][9] : "N"
      });
    }
    return tips.reverse(); // 최신순
  } catch(e) { return []; }
}



function saveWisdomTip(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("지혜의_보물고") || ss.insertSheet("지혜의_보물고");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["날짜", "제목", "내용", "카테고리", "작성자", "조회수", "공감수", "깨달음수", "사진주소", "기본게시물"]);
      sheet.setFrozenRows(1);
    }
    
    var rawImage = data.image || "";
    var isBase64 = (rawImage.indexOf("data:") === 0 || rawImage.length > 1000);
    var cleanImage = isBase64
      ? saveBase64ImageToDrive(rawImage, "wisdom_img_" + Date.now() + ".png")
      : cleanPollinationsApiKey(rawImage);
      
    sheet.appendRow([
      new Date(), 
      data.title, 
      data.content, 
      data.category, 
      "길드마스터", 
      0, // 조회수
      0, // 공감수
      0, // 깨달음수
      cleanImage, // 사진주소
      data.isDefault || "N" // 기본게시물여부 (Y/N)
    ]);
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}



function updateWisdomTip(rowIdx, data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("지혜의_보물고");
    if (!sheet) return { success: false, error: "지혜의_보물고 시트가 존재하지 않습니다." };
    
    var row = parseInt(rowIdx);
    if (isNaN(row) || row < 2 || row > sheet.getLastRow()) {
      return { success: false, error: "유효하지 않은 행 식별자입니다: " + rowIdx };
    }
    
    // 시트 컬럼 구조:
    // A: 날짜(1), B: 제목(2), C: 내용(3), D: 카테고리(4), E: 작성자(5)
    // I: 사진주소(9), J: 기본게시물(10)
    var rawImage = data.image || "";
    var isBase64 = (rawImage.indexOf("data:") === 0 || rawImage.length > 1000);
    var cleanImage = isBase64
      ? saveBase64ImageToDrive(rawImage, "wisdom_img_" + Date.now() + ".png")
      : cleanPollinationsApiKey(rawImage);
      
    sheet.getRange(row, 2).setValue(data.title);
    sheet.getRange(row, 3).setValue(data.content);
    sheet.getRange(row, 4).setValue(data.category);
    sheet.getRange(row, 9).setValue(cleanImage);
    sheet.getRange(row, 10).setValue(data.isDefault || "N");
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}



function createSurpriseQuest(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("마을_이벤트") || ss.insertSheet("마을_이벤트");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["제목", "내용", "배율", "마감시간", "상태", "유형"]);
    }
    
    var now = new Date();
    var endTime = new Date(now.getTime() + (data.durationMinutes * 60000));
    
    sheet.appendRow([
      data.title, 
      data.description, 
      data.multiplier || 1, 
      endTime, 
      "ACTIVE", 
      data.type || "NEW"
    ]);
    
    // 전령의 기둥에도 자동 선포
    updatePillarNotice({
      content: `⚡ [돌발 계시] ${data.title}! (${data.multiplier}배 보상 / ${data.durationMinutes}분 한정)`,
      category: "QUEST"
    });
    
    return { success: true, endTime: endTime };
  } catch(e) { return { success: false, error: e.toString() }; }
}



function checkAndCreateQuestRegistrySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("돌발퀘스트_목록");
  var headers = ["ID", "시행일", "유형", "퀘스트명", "설명", "만료일", "전화번호", "상태", "점수", "인증방식"];
  if (!sheet) {
    sheet = ss.insertSheet("돌발퀘스트_목록");
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
    sheet.getRange("1:1").setBackground("#e3faf2").setFontWeight("bold").setHorizontalAlignment("center");
    Logger.log("✅ '돌발퀘스트_목록' 시트 신규 생성 및 10열 헤더 배치 완료!");
  } else {
    // 🩹 기존 시트의 헤더가 유실되었거나 8열까지만 있는 경우 자가치유(Self-healing) 실행
    var lastCol = sheet.getLastColumn();
    var currentHeaders = sheet.getRange(1, 1, 1, Math.max(lastCol, 10)).getDisplayValues()[0];
    var needsRepair = false;
    for (var c = 0; c < headers.length; c++) {
      if (String(currentHeaders[c] || "").trim() !== headers[c]) {
        needsRepair = true;
        break;
      }
    }
    if (needsRepair) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange("1:1").setBackground("#e3faf2").setFontWeight("bold").setHorizontalAlignment("center");
      Logger.log("🩹 '돌발퀘스트_목록' 시트 헤더 자가치유(10열 확장) 완료!");
    }
  }
  return sheet;
}



function triggerGlycogenQuest(phone, name) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndCreateQuestRegistrySheet();
    var cleanPhone = String(phone).replace(/[^0-9]/g, "");
    
    var data = sheet.getDataRange().getDisplayValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === "글리코겐" && String(data[i][6]).replace(/[^0-9]/g, "") === cleanPhone && data[i][7] === "진행중") {
        Logger.log("이미 진행중인 글리코겐 퀘스트가 있습니다.");
        return;
      }
    }
    
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var deadline = new Date(now.getTime() + 72 * 60 * 60 * 1000); // 72 hours
    var deadlineStr = Utilities.formatDate(deadline, "GMT+9", "yyyy-MM-dd HH:mm:ss");
    
    var nextId = "GLY_" + Date.now();
    
    sheet.appendRow([
      nextId,
      todayStr,
      "글리코겐",
      "🔥 글리코겐 클리어 퀘스트",
      "치팅으로 인한 간과 근육의 글리코겐이 포착되었습니다! 3일 연속 출석하여 방패를 채우세요!",
      deadlineStr,
      "'" + cleanPhone,
      "진행중"
    ]);

    // 🛡️ 글리코겐 방패 최초 발동 쪽지 즉시 발송
    try {
      var mName = name;
      if (!mName) mName = "회원";
      mName = String(mName).replace(/\d{4}$/, ""); // 이름 뒤 숫자 제거

      sendPersonalNotification(
        cleanPhone,
        "방어",
        "🚨 글리코겐 방패 긴급 발동! 3일 연속 출석 미션! 🛡️",
        mName + "님, 저녁을 무겁게 드셨거나 치팅이 감지되어 글리코겐 방패가 긴급 발동되었습니다! 🚨\n\n글리코겐이 지방으로 축적되기 전에 3일 안에 꼭 출석하셔서 운동으로 완전히 연소시켜야 합니다! 내일부터 3일간 꼭 출석하셔서 요요 방어 방패를 완성해보세요. 화이팅! 🏃‍♀️"
      );
    } catch (notiErr) {
      Logger.log("글리코겐 최초 발동 알림 생성 실패: " + notiErr.toString());
    }

    Logger.log("글리코겐 퀘스트 성공적으로 발동됨!");
  } catch(e) {
    Logger.log("triggerGlycogenQuest 에러: " + e.toString());
  }
}



/**
 * 0. 마을_공지 시트의 헤더 무결성 보장 및 강제 초기화 헬퍼 함수
 */
function checkAndInitNoticeSheet(ss) {
  var sheet = ss.getSheetByName("마을_공지");
  if (!sheet) {
    sheet = ss.insertSheet("마을_공지");
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["작성일", "카테고리", "제목", "내용", "활성화"]);
    var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
    sheet.appendRow([
      todayStr,
      "선포",
      "노형 빌리지 대광장 리뉴얼 선포!",
      "노형 빌리지의 광장과 웰니스 센터가 새롭게 태어났습니다. 이제 마을 공지 배너를 누르면 이 황금 두루마리를 통해 언제든 마을의 중대사 역사와 공지 기록을 확인할 수 있습니다. ✨",
      "TRUE"
    ]);
  } else {
    // 1행 첫 번째 셀이 헤더가 아닌 일반 데이터라면 1행에 헤더를 강제 삽입하고 기존 데이터는 아래로 밀어 복원!
    var firstCellVal = String(sheet.getRange(1, 1).getValue()).trim();
    if (firstCellVal !== "작성일" && firstCellVal !== "선포일") {
      sheet.insertRowBefore(1);
      sheet.getRange(1, 1, 1, 5).setValues([["작성일", "카테고리", "제목", "내용", "활성화"]]);
    }
  }
  return sheet;
}



/**
 * 1. 실시간 전령의 기둥(공지) 내용 조회 (마을_공지 시트 단일 소스 최적화)
 */
function getPillarNotice() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndInitNoticeSheet(ss);
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [{ title: "오늘도 건강한 하루 되세요! 📢", content: "" }];
    
    var activeNotices = [];
    // 모든 활성화된 공지 가져오기 (가장 최신 공지부터 역순으로 수집)
    for (var i = data.length - 1; i >= 1; i--) {
      var activeVal = String(data[i][4]).toUpperCase().trim();
      if (activeVal === "TRUE" || activeVal === "ACTIVATED") {
        activeNotices.push({ 
          title: data[i][2],   // C열 제목 (index 2)
          content: data[i][3]  // D열 내용 (index 3)
        });
      }
    }
    
    if (activeNotices.length === 0) {
      return [{ title: "오늘도 건강한 하루 되세요! 📢", content: "" }];
    }
    return activeNotices;
  } catch (e) {
    return [{ title: "오늘도 건강한 하루 되세요! 📢", content: "" }];
  }
}



/**
 * 2. 실시간 전령의 기둥(공지) 내용 업데이트 (마을_공지 시트 단일 소스 최적화)
 */
function updatePillarNotice(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndInitNoticeSheet(ss);
    
    var todayStr = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd");
    var category = payload.category || "공지";
    var title = payload.title || "";
    var content = payload.content || "";
    
    sheet.appendRow([todayStr, category, title, content, "TRUE"]);
    
    // ⚡ [공지 캐시 초기화] 새 공지가 작성되면 서버 캐시를 즉각 비워 실시간 반영 보장!
    var cache = CacheService.getScriptCache();
    cache.remove("v58_village_notices");
    cache.remove("v58_pillar_notices");
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}



/**
 * 4. 오늘 하루 하사된 총 경험치 집계
 */
function getDailyAwardedExperience() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("일일_활동_기록");
    if (!sheet) return 0;
    var data = sheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    var totalExp = 0;
    for (var i = 1; i < data.length; i++) {
      var dateRaw = data[i][0];
      var dateStr = (dateRaw instanceof Date) ? Utilities.formatDate(dateRaw, "GMT+9", "yyyy-MM-dd") : String(dateRaw);
      if (dateStr.indexOf(todayStr) > -1) {
        totalExp += (parseInt(data[i][9]) || 0); // J열: 총점
      }
    }
    return totalExp;
  } catch (e) {
    return 0;
  }
}


function checkAndAwardWeeklyAttendanceBonus(phoneStr, memberName) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("일일_활동_기록");
    if (!sheet) return;
    
    var now = new Date();
    // [v46.43] 주간: 매주 월요일 00:00:00 기점 (일요일 자정 마감)
    var day = now.getDay(); // 0(일)~6(토)
    var diffToMon = (day === 0) ? 6 : (day - 1); 
    var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - diffToMon);
    startOfWeek.setHours(0, 0, 0, 0);
    
    var data = sheet.getDataRange().getValues();
    var attendDaysThisWeek = 0;
    
    // 보너스 수령 여부 플래그
    var hasBonus3 = false;
    var hasBonus4 = false;
    var hasBonus5 = false;
    var hasBonus6 = false;
    
    for (var i = 1; i < data.length; i++) {
      var rowDate = new Date(data[i][0]);
      if (rowDate >= startOfWeek) {
        var rowPhone = String(data[i][1]).replace(/[^0-9]/g, "");
        if (rowPhone === phoneStr) {
          // 센터방문 수행 실적이 있는 날 카운트
          var visitScore = Number(data[i][3]) || 0;
          if (visitScore > 0) {
            attendDaysThisWeek++;
          }
          // 보너스 항목 수령 여부 확인
          var details = String(data[i][8] || "");
          if (details.indexOf("주간 3회 출석 보너스") > -1) hasBonus3 = true;
          if (details.indexOf("주간 4회 출석 보너스") > -1) hasBonus4 = true;
          if (details.indexOf("주간 5회 출석 보너스") > -1) hasBonus5 = true;
          if (details.indexOf("주간 6회 출석 보너스") > -1) hasBonus6 = true;
        }
      }
    }
    
    var friendlyName = getFriendlyName(memberName);
    
    // 3일 출석: +100 EXP 하사 (누적 100)
    if (attendDaysThisWeek >= 3 && !hasBonus3) {
      recordActivityLog({
        phone: phoneStr,
        name: memberName,
        type: "보너스",
        item: "주간 3회 출석 보너스",
        action: "달성",
        score: 100,
        statType: "health"
      });
      Logger.log("🎉 [" + memberName + "] 주간 3회 출석 달성! 체력 보너스 +100 EXP 하사 완료!");
      sendWeeklyBonusNotification(phoneStr, memberName, 3);
      hasBonus3 = true;
    }
    // 4일 출석: +100 EXP 추가 하사 (누적 200)
    if (attendDaysThisWeek >= 4 && !hasBonus4) {
      recordActivityLog({
        phone: phoneStr,
        name: memberName,
        type: "보너스",
        item: "주간 4회 출석 보너스",
        action: "달성",
        score: 100,
        statType: "health"
      });
      Logger.log("🎉 [" + memberName + "] 주간 4회 출석 달성! 체력 보너스 +100 EXP 추가 하사 완료!");
      sendWeeklyBonusNotification(phoneStr, memberName, 4);
      hasBonus4 = true;
    }
    // 5일 출석: +100 EXP 추가 하사 (누적 300)
    if (attendDaysThisWeek >= 5 && !hasBonus5) {
      recordActivityLog({
        phone: phoneStr,
        name: memberName,
        type: "보너스",
        item: "주간 5회 출석 보너스",
        action: "달성",
        score: 100,
        statType: "health"
      });
      Logger.log("🎉 [" + memberName + "] 주간 5회 출석 달성! 체력 보너스 +100 EXP 추가 하사 완료!");
      sendWeeklyBonusNotification(phoneStr, memberName, 5);
      hasBonus5 = true;
    }
    // 6일 출석: +100 EXP 추가 하사 (누적 400 EXP!)
    if (attendDaysThisWeek >= 6 && !hasBonus6) {
      recordActivityLog({
        phone: phoneStr,
        name: memberName,
        type: "보너스",
        item: "주간 6회 출석 보너스",
        action: "달성",
        score: 100,
        statType: "health"
      });
      Logger.log("🎉 [" + memberName + "] 주간 6회 출석 달성! 체력 보너스 +100 EXP 추가 하사 완료!");
      sendWeeklyBonusNotification(phoneStr, memberName, 6);
      hasBonus6 = true;
    }
  } catch (e) {
    Logger.log("🚨 주간 출석 보너스 정산 오류: " + e.toString());
  }
}



/**
 * [v60.3] 체력 보너스 달성 시 일반문자(SMS)를 발송하지 않고, 인앱 쪽지(Personal Notification)로 안전하게 즉시 발송
 */
function sendWeeklyBonusNotification(phoneStr, memberName, count) {
  try {
    var friendlyName = getFriendlyName(memberName);
    var title = "주간 " + count + "회 출석 보너스 달성! 🎉";
    var content = friendlyName + "님, 축하합니다! 금주 " + count + "회째 출석 마일스톤을 달성하여 체력보너스 추가 +100 EXP가 전격 지급되었습니다! 오늘도 신나고 건강하게 점핑! 🔥🤸‍♂️";
    
    if (count === 4) {
      content = friendlyName + "님, 정말 대단하십니다! 금주 " + count + "회째 출석 마일스톤을 달성하여 체력보너스 추가 +100 EXP가 전격 지급되었습니다! 꾸준함이 모여 기적을 만듭니다. 화이팅! 🏆✨";
    } else if (count === 5) {
      content = friendlyName + "님, 당신은 진정한 웰니스 리더! 금주 " + count + "회째 출석 마일스톤을 달성하여 체력보너스 추가 +100 EXP가 전격 지급되었습니다! 오늘 하루도 상쾌하고 에너제틱하게! 💪🔥";
    } else if (count >= 6) {
      title = "주간 6회 출석 그랜드슬램 달성! 👑";
      content = friendlyName + "님, 경이로운 출석 질주! 금주 " + count + "회째 출석 그랜드슬램을 달성하여 체력보너스 추가 +100 EXP가 전격 지급되었습니다! 건강한 습관의 끝판왕, 웰니스 코치가 진심으로 축하드립니다! ❤️🤸‍♂️";
    }
    
    sendPersonalNotification(phoneStr, "welcome", title, content);
    Logger.log("📢 [" + memberName + "] 체력보너스 인앱 쪽지 안전 전송 완료 (주간 " + count + "회)");
  } catch (e) {
    Logger.log("🚨 체력보너스 인앱 쪽지 전송 오류: " + e.toString());
  }
}



/**
 * [트리거 호환 안전 패치 v60.41]
 * 혹시 구글 Apps Script 온라인 트리거에 구버전 함수명(sendWeeklyBonusSms)이 수동 등록되어 돌고 있을 경우를 대비하여,
 * 실행 차단 에러를 예방하고 인앱 쪽지 발송 신규 로직으로 안전하게 토스해 주는 하위 호환용 연결 함수입니다.
 */
function sendWeeklyBonusSms(phoneStr, memberName, count) {
  Logger.log("⚠️ [호환 엔진] 구버전 트리거 함수명(sendWeeklyBonusSms) 감지 -> 신규 인앱 쪽지 엔진으로 자동 이관합니다.");
  sendWeeklyBonusNotification(phoneStr, memberName, count);
}


function getGeminiApiKey() {
  try {
    // 1단계: "환경설정" 시트의 B2:B6 영역에서 복수 키들을 실시간으로 수집합니다.
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("환경설정");
    var sheetKeys = [];
    if (sheet) {
      var values = sheet.getRange("B2:B6").getValues();
      for (var i = 0; i < values.length; i++) {
        var k = String(values[i][0] || "").trim();
        // 옛날 차단 화석 키나 헤더 등 비정상 값 필터링
        if (k && k !== "API_KEY" && !k.startsWith("AIzaSyAzX")) {
          sheetKeys.push(k);
        }
      }
    }

    var cachedKey = String(PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "").trim();

    // 만약 시트에서 키를 하나도 못 가져왔을 때, 기존 캐시를 최후의 보루로 활용
    if (sheetKeys.length === 0) {
      if (cachedKey && !cachedKey.startsWith("AIzaSyAzX")) {
        sheetKeys.push(cachedKey);
      }
    } else {
      // 캐시 갱신 (1순위 키 기준)
      if (sheetKeys[0] !== cachedKey) {
        PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", sheetKeys[0]);
      }
    }

    var primaryKey = sheetKeys.length > 0 ? sheetKeys[0] : "";

    var pKeyRes = getPollinationsApiKey();
    var pKey = pKeyRes.success ? pKeyRes.key : "";

    return { 
      success: true, 
      key: primaryKey,          // 하위 호환성 유지용 단일 키
      keys: sheetKeys,          // 실시간 피스톤 로테이션용 복수 키 배열
      pollinationsKey: pKey     // Pollinations API Key
    };
  } catch (e) {
    Logger.log("🚨 getGeminiApiKey 오류: " + e.toString());
    return { success: false, error: e.toString() };
  }
}



/**
 * [v67.7] Pollinations AI API Key 로드 함수 (402 결제 오류 예방용)
 */
function getPollinationsApiKey() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("환경설정");
    var key = "";
    if (sheet) {
      // B7 셀에서 Pollinations Key를 가져옵니다.
      key = String(sheet.getRange("B7").getValue() || "").trim();
    }
    
    var cachedKey = String(PropertiesService.getScriptProperties().getProperty("POLLINATIONS_API_KEY") || "").trim();
    if (!key) {
      // B7 셀이 비어있으면 캐시된 키를 명시적으로 삭제하고 빈 값을 사용합니다.
      if (cachedKey) {
        PropertiesService.getScriptProperties().deleteProperty("POLLINATIONS_API_KEY");
      }
      key = "";
    } else if (key !== cachedKey) {
      PropertiesService.getScriptProperties().setProperty("POLLINATIONS_API_KEY", key);
    }
    
    return { success: true, key: key };
  } catch (e) {
    Logger.log("🚨 getPollinationsApiKey 오류: " + e.toString());
    return { success: false, error: e.toString() };
  }
}



/**
 * [v67.8] Pollinations AI 이미지 URL에 API Key를 실시간으로 결합해 주는 헬퍼 함수
 */
function appendPollinationsApiKey(imgUrl, pKey) {
  if (!imgUrl || !pKey) return imgUrl;
  var hasKey = imgUrl.indexOf("key=") > -1;
  if (hasKey) return imgUrl;
  
  var isLegacy = imgUrl.indexOf("image.pollinations.ai/prompt/") > -1;
  var isNew = imgUrl.indexOf("gen.pollinations.ai/image/") > -1;
  
  if (isNew) {
    return imgUrl + (imgUrl.indexOf("?") > -1 ? "&" : "?") + "key=" + pKey;
  }
  
  if (isLegacy) {
    // sk_ 비밀키는 브라우저 이미지 요청에 사용되면 402 에러를 유발하므로 제외
    if (pKey.startsWith("pk_")) {
      return imgUrl + (imgUrl.indexOf("?") > -1 ? "&" : "?") + "key=" + pKey;
    }
  }
  
  // 기타 일반적인 경우
  if ((imgUrl.indexOf("pollinations.ai") > -1) && pKey.startsWith("pk_")) {
    return imgUrl + (imgUrl.indexOf("?") > -1 ? "&" : "?") + "key=" + pKey;
  }
  
  return imgUrl;
}



/**
 * [v67.8] Pollinations AI 이미지 URL에서 API Key 매개변수를 깨끗하게 제거하는 헬퍼 함수
 */
function cleanPollinationsApiKey(imgUrl) {
  if (!imgUrl) return "";
  if (imgUrl.indexOf("pollinations.ai") > -1) {
    var cleaned = imgUrl.replace(/[?&]key=[^&]*/g, "");
    if (cleaned.endsWith("?") || cleaned.endsWith("&")) {
      cleaned = cleaned.substring(0, cleaned.length - 1);
    }
    return cleaned;
  }
  return imgUrl;
}



/**
 * [v68.0] 지식창고 이미지 구글 드라이브 전용 보관 폴더 생성 및 설정
 */
function getOrCreateWisdomFolder() {
  var folderName = "GenieWorld_Wisdom_Images";
  var scriptProps = PropertiesService.getScriptProperties();
  var cachedFolderId = scriptProps.getProperty("wisdom_folder_id");
  var folder;
  if (cachedFolderId) {
    try {
      folder = DriveApp.getFolderById(cachedFolderId);
    } catch (e) {
      Logger.log("캐싱된 폴더 조회 실패, 새로 조회합니다: " + e.toString());
    }
  }
  if (!folder) {
    var folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }
    // 부모 폴더에 권한 설정 (한 번만 호출하므로 매우 빠름, 하위 파일 상속)
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    scriptProps.setProperty("wisdom_folder_id", folder.getId());
  }
  return folder;
}



/**
 * [v68.0] Base64 이미지 데이터를 구글 드라이브에 저장하고 direct URL 반환
 */
function saveBase64ImageToDrive(base64Data, fileName) {
  try {
    var folder = getOrCreateWisdomFolder();
    var cleanBase64 = base64Data;
    var contentType = "image/png";
    
    if (cleanBase64.indexOf(",") > -1) {
      var parts = cleanBase64.split(",");
      var meta = parts[0];
      cleanBase64 = parts[1];
      if (meta.indexOf("image/jpeg") > -1) contentType = "image/jpeg";
      else if (meta.indexOf("image/jpg") > -1) contentType = "image/jpeg";
      else if (meta.indexOf("image/gif") > -1) contentType = "image/gif";
    }
    
    cleanBase64 = cleanBase64.replace(/\s/g, '');
    var decoded = Utilities.base64Decode(cleanBase64);
    var blob = Utilities.newBlob(decoded, contentType, fileName);
    
    var file = folder.createFile(blob);
    return "https://lh3.googleusercontent.com/d/" + file.getId();
  } catch (e) {
    Logger.log("🚨 saveBase64ImageToDrive 에러: " + e.toString());
    return "";
  }
}



/**
 * [v68.0] 프론트엔드에서 전송받은 특정 팁 ID의 Base64 데이터를 드라이브에 저장하고 시트 업데이트
 */
function migrateSingleImage(tipId, base64Data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("지혜의_보물고");
    if (!sheet) return { success: false, error: "지혜의_보물고 시트가 존재하지 않습니다." };
    
    var row = parseInt(tipId);
    if (isNaN(row) || row < 2 || row > sheet.getLastRow()) {
      return { success: false, error: "유효하지 않은 행 번호입니다: " + tipId };
    }
    
    var timestamp = Date.now();
    var driveUrl = saveBase64ImageToDrive(base64Data, "wisdom_migrated_" + tipId + "_" + timestamp + ".png");
    
    if (driveUrl && driveUrl.indexOf("googleusercontent.com") > -1) {
      sheet.getRange(row, 9).setValue(driveUrl); // I열 (9열)
      return { success: true, url: driveUrl };
    } else {
      return { success: false, error: "드라이브 저장 후 링크 생성 실패" };
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}



function setGeminiApiKey(newKey) {
  try {
    var trimmed = String(newKey || "").trim();
    if (!trimmed) {
      PropertiesService.getScriptProperties().deleteProperty("GEMINI_API_KEY");
      return { success: true, message: "API 키가 안전하게 제거되었습니다." };
    }
    
    // 1단계: Properties에 안전하게 암호화 저장
    PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", trimmed);
    
    // 2단계: 이장님 확인용으로 "환경설정" 시트 생성 및 쓰기
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("환경설정");
    if (!sheet) {
      sheet = ss.insertSheet("환경설정");
      sheet.getRange("A1").setValue("설명");
      sheet.getRange("A2").setValue("Gemini AI API Key (소스코드 유출 방지용 금고)");
      sheet.getRange("B1").setValue("API_KEY");
    }
    sheet.getRange("B2").setValue(trimmed);
    
    return { success: true, message: "구글 스프레드시트 금고에 API Key가 안전하게 저장되었습니다!" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}



/**
 * 🤖 [v64.30] 지니 웰니스 AI 비서 (Gemini 1.5 Flash) 백엔드 API 직접 호출 엔진
 */
function callGeminiBackendWithDetails(prompt, systemInstruction) {
  try {
    var apiKeyRes = getGeminiApiKey();
    if (!apiKeyRes.success || !apiKeyRes.keys || apiKeyRes.keys.length === 0) {
      return { success: false, error: "API Key가 로드되지 않았습니다." };
    }
    var keys = apiKeyRes.keys;
    
    // 🌟 원장님의 명설계: 안정적인 멀티 티어 모델 우회(폴백) 리스트 (Gemini 2.5 Flash 시리즈 100% 최적화)
    var fallbackModels = [
      "gemini-2.5-flash",
      "gemini-2.5-flash-lite",
      "gemini-2.0-flash",
      "gemini-2.0-flash-lite"
    ];
    
    var payload = {
      "contents": [
        {
          "role": "user",
          "parts": [{"text": prompt}]
        }
      ],
      "generationConfig": {
        "temperature": 0.7
      }
    };
    
    if (systemInstruction) {
      payload["systemInstruction"] = {
        "parts": [{"text": systemInstruction}]
      };
    }
    
    var options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };
    
    var errorLogs = [];
    
    // 🔑 바깥 루프: API Key 로테이션 (최대 5개 복수 키)
    for (var k = 0; k < keys.length; k++) {
      var apiKey = keys[k];
      var keySnippet = apiKey.substring(0, 6) + "..." + apiKey.substring(apiKey.length - 4);
      Logger.log("🔑 [백엔드 AI] API Key 시도 [" + (k + 1) + "/" + keys.length + "]: " + keySnippet);
      
      var keyFailed = false;
      
      // 🤖 안쪽 루프: 모델 폴백 (gemini-2.5-flash -> gemini-2.5-flash-lite)
      for (var i = 0; i < fallbackModels.length; i++) {
        var modelName = fallbackModels[i];
        var url = "https://generativelanguage.googleapis.com/v1beta/models/" + modelName + ":generateContent?key=" + apiKey;
        
        try {
          Logger.log("🔮 [백엔드 AI] 모델 시도 [" + (i + 1) + "단계]: " + modelName + "...");
          var response = UrlFetchApp.fetch(url, options);
          var responseCode = response.getResponseCode();
          var responseText = response.getContentText();
          
          if (responseCode === 200) {
            var json = JSON.parse(responseText);
            if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts[0]) {
              Logger.log("✨ [백엔드 AI] 성공: [" + modelName + "] 모델이 완수했습니다. (Key: " + keySnippet + ")");
              return { success: true, text: json.candidates[0].content.parts[0].text.trim() };
            }
          }
          
          errorLogs.push("Key [" + keySnippet + "] Model [" + modelName + "] 실패: HTTP " + responseCode + " - " + responseText);
          Logger.log("⚠️ [백엔드 AI] " + modelName + " 실패: HTTP " + responseCode);
  
          // 💡 [v66.10] 429, 400뿐만 아니라 503, 502 등 200이 아닌 모든 이상 상태 코드 발생 시 즉시 다음 키로 로테이션 가동!
          if (responseCode !== 200) {
            Logger.log("🚨 [서버/모델 장애 감지] API Key [" + keySnippet + "]에서 HTTP " + responseCode + " 감지. 다음 우회 모델로 즉시 바통 터치합니다.");
            // 🌟 [원장님의 조언 수용] break 제거! 모델 루프를 끝까지 순환하며 예비 모델을 모두 찔러보도록 유도
            continue;
          }
          
          Utilities.sleep(500);
        } catch (e) {
          errorLogs.push("Key [" + keySnippet + "] Model [" + modelName + "] 에러: " + e.toString());
          Logger.log("🚨 [백엔드 AI] " + modelName + " 에러: " + e.toString());
          Utilities.sleep(500);
        }
      }
      
      // 만약 429/400이 아니고 다른 특별한 에러가 나더라도, 바깥 루프가 계속 돌며 다음 키를 시도할 수 있도록 지원합니다.
    }
    
    return { success: false, error: "모든 등록된 API Key 및 모델 호출에 실패했습니다.\n\n상세 정보:\n" + errorLogs.join("\n\n") };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}



function callGeminiBackend(prompt, systemInstruction) {
  var res = callGeminiBackendWithDetails(prompt, systemInstruction);
  if (res.success) {
    return res.text;
  }
  Logger.log("🚨 Gemini API 호출 에러: " + res.error);
  return null;
}


/**
 * [v48.0] 실시간 제주의 실제 기상 상황을 조회하여 가상 기후와 연동해주는 마법 에이전트 엔진
 * Open-Meteo API를 사용하며, 서버 부하와 속도 유지를 위해 1시간 동안 조회 결과를 캐싱합니다.
 */
function getJejuRealtimeWeather() {
  var cache = CacheService.getScriptCache();
  var cachedWeather = cache.get("jeju_realtime_weather");
  if (cachedWeather) {
    try {
      return JSON.parse(cachedWeather);
    } catch (e) {
      cache.remove("jeju_realtime_weather");
    }
  }

  // WMO 날씨 코드가 아니거나 API 실패 시 사용할 계절 기반 기본 날씨 분석
  var currentMonth = new Date().getMonth() + 1; // 1-12월
  var defaultSeasonalWeather = "sun";
  if (currentMonth >= 3 && currentMonth <= 5) {
    defaultSeasonalWeather = "blossom"; // 봄: 벚꽃
  } else if (currentMonth >= 9 && currentMonth <= 11) {
    defaultSeasonalWeather = "leaves";  // 가을: 낙엽
  }

  try {
    // 제주시 노형동(클럽 위치) 기준 위도(33.4781), 경도(126.4755) API 호출
    var url = "https://api.open-meteo.com/v1/forecast?latitude=33.4781&longitude=126.4755&current_weather=true";
    var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      var json = JSON.parse(response.getContentText());
      var currentWeather = json.current_weather;
      var weatherCode = currentWeather.weathercode;
      
      var windSpeed = currentWeather.windspeed; // km/h
      var windSpeedMs = Number((windSpeed / 3.6).toFixed(1)); // m/s
      
      var weather = defaultSeasonalWeather;
      // WMO Weather Codes 번역:
      // - 비(Rain/Drizzle): 51, 53, 55, 61, 63, 65, 80, 81, 82, 95, 96, 99
      // - 눈(Snow): 71, 73, 75, 85, 86
      // - 맑음/흐림(Seasonal Blossom/Leaves/Sun): 그 외 코드
      if ([51, 53, 55, 61, 63, 65, 80, 81, 82, 95, 96, 99].indexOf(weatherCode) > -1) {
        weather = "rain";
      } else if ([71, 73, 75, 85, 86].indexOf(weatherCode) > -1) {
        weather = "snow";
      } else {
        weather = defaultSeasonalWeather;
      }

      var result = { weather: weather, temp: currentWeather.temperature, wind: windSpeedMs };
      // 30분(1800초) 캐싱 최적화 반영
      cache.put("jeju_realtime_weather", JSON.stringify(result), 1800);
      return result;
    }
  } catch (e) {
    Logger.log("제주 기상 데이터 동기화 실패: " + e.toString());
  }
  return { weather: defaultSeasonalWeather, temp: 20, wind: 0 };
}



function getVillageSettings() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("global_village_settings");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {
      cache.remove("global_village_settings");
    }
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("마을_설정");
    if (!sheet) {
      sheet = ss.insertSheet("마을_설정");
      sheet.appendRow(["키", "값", "설명"]);
      sheet.appendRow(["weather", "sun", "마을 날씨 (sun, rain, snow, blossom, leaves, auto)"]);
      sheet.appendRow(["bgmEnabled", "false", "배경음악 활성화 여부 (true, false)"]);
      sheet.appendRow(["bgmUrl", "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3", "배경음악 음원 주소"]);
      sheet.appendRow(["bgm_sun", "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3", "맑음 테마 배경음악"]);
      sheet.appendRow(["bgm_rain", "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-8.mp3", "비 테마 배경음악"]);
      sheet.appendRow(["bgm_snow", "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-10.mp3", "눈 테마 배경음악"]);
      sheet.appendRow(["bgm_blossom", "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-3.mp3", "벚꽃 테마 배경음악"]);
      sheet.appendRow(["bgm_leaves", "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-5.mp3", "낙엽 테마 배경음악"]);
      sheet.setFrozenRows(1);
    }
    
    var data = sheet.getDataRange().getDisplayValues();
    var settings = {};
    var keysInSheet = {};
    for (var i = 1; i < data.length; i++) {
      settings[data[i][0]] = data[i][1];
      keysInSheet[data[i][0]] = i + 1;
    }

    // [v48.0] 만약 기존 마을_설정 시트에 기후별 BGM 키가 누락되어 있다면 자동 복구/추가
    var defaultBgmMap = {
      "bgm_sun": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3",
      "bgm_rain": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-8.mp3",
      "bgm_snow": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-10.mp3",
      "bgm_blossom": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-3.mp3",
      "bgm_leaves": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-5.mp3",
      "bgmForceOverride": "false",
      "bgmForceUrl": "",
      "magicEnabled": "true"
    };

    var addedNewRow = false;
    for (var key in defaultBgmMap) {
      if (!keysInSheet[key]) {
        sheet.appendRow([key, defaultBgmMap[key], "기후별 커스텀 배경음악 및 고정 설정"]);
        settings[key] = defaultBgmMap[key];
        addedNewRow = true;
      }
    }
    if (addedNewRow) {
      data = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < data.length; i++) {
        settings[data[i][0]] = data[i][1];
      }
    }

    // [v48.0] 실시간 제주시 노형동 기상싱크 모드 발동 처리
    if (settings.weather === "auto") {
      var jeju = getJejuRealtimeWeather();
      settings.resolvedWeather = jeju.weather;
      settings.realJejuTemp = jeju.temp; // 현재 제주의 실시간 온도 정보도 화면단에 함께 주입!
      settings.realJejuWind = jeju.wind; // 현재 제주의 실시간 풍속 정보(m/s) 주입!
    } else {
      settings.resolvedWeather = settings.weather;
    }

    // [원장님 직관적 설계 반영] 강제 고정 모드와 기후별 연동 분기 처리 (Google Sheets 대소문자 TRUE/FALSE 극복)
    var isBgmEnabled = settings.bgmEnabled && settings.bgmEnabled.toString().toLowerCase() === "true";
    if (isBgmEnabled) {
      var isForce = settings.bgmForceOverride && settings.bgmForceOverride.toString().toLowerCase() === "true";
      var forceUrl = settings.bgmForceUrl || "";
      if (isForce && forceUrl.trim() !== "") {
        // 1. 강제 고정 모드가 켜져 있고 음악 주소가 있으면 무조건 강제 고정곡 재생!
        settings.bgmUrl = forceUrl;
      } else {
        // 2. 그 외에는 100% 날씨와 한 세트로 묶여서 재생!
        var weatherKey = "bgm_" + (settings.resolvedWeather || "sun");
        settings.bgmUrl = settings[weatherKey] || settings["bgm_sun"] || defaultBgmMap["bgm_sun"];
      }
    }
    
    // [v48.0] 불러올 때도 기존에 시트에 박혀있는 모든 Suno 단축/공유 링크를 오디오 다이렉트 주소로 초고속 해독!
    // 오직 실제 BGM URL 주소 값을 가진 필드들만 엄격히 발라내어 해독 연산 수행! (불리언/설정 플래그 오작동 방지)
    // 💡 [마법의 백포트 보정 엔진] 원장님이 구글 시트에 직접 날것의 Suno 주소를 복사해 넣더라도, 로드 시 알아서 해독하여 시트 셀에 역저장(자동 변환)합니다!
    var bgmKeys = ["bgm_sun", "bgm_rain", "bgm_snow", "bgm_blossom", "bgm_leaves", "bgmForceUrl", "bgmUrl"];
    bgmKeys.forEach(function(key) {
      if (settings[key]) {
        var originalVal = String(settings[key]).trim();
        var resolvedVal = resolveSunoUrl(originalVal);
        if (originalVal !== resolvedVal) {
          settings[key] = resolvedVal;
          // 만약 해석 전/후 주소가 달라졌다면 시트에도 자동으로 변환된 값을 역저장하여 영구 업데이트!
          if (keysInSheet[key]) {
            sheet.getRange(keysInSheet[key], 2).setValue(resolvedVal);
          }
        }
      }
    });
    
    // 5분 동안 설정 캐시 저장 (update 시 즉시 날라감)
    cache.put("global_village_settings", JSON.stringify(settings), 300);
    return settings;
  } catch (e) {
    return { weather: "sun", bgmEnabled: "false", bgmUrl: "" };
  }
}



function updateVillageSettings(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("마을_설정") || ss.insertSheet("마을_설정");
    
    var settingsMap = {
      "weather": payload.weather || "sun",
      "bgmEnabled": String(payload.bgmEnabled || "false"),
      "bgmForceOverride": String(payload.bgmForceOverride || "false"),
      "bgmForceUrl": payload.bgmForceUrl ? resolveSunoUrl(payload.bgmForceUrl.trim()) : "",
      "magicEnabled": String(payload.magicEnabled !== undefined ? payload.magicEnabled : "true")
    };

    // [v48.0] 기후별 커스텀 BGM 키 동적 수신 및 매핑 저장
    var customKeys = ["bgm_sun", "bgm_rain", "bgm_snow", "bgm_blossom", "bgm_leaves"];
    customKeys.forEach(function(k) {
      if (payload[k] !== undefined) {
        settingsMap[k] = resolveSunoUrl(payload[k]);
      }
    });
    
    var data = sheet.getDataRange().getValues();
    var keysInSheet = {};
    for (var i = 1; i < data.length; i++) {
      keysInSheet[data[i][0]] = i + 1;
    }
    
    for (var k in settingsMap) {
      if (keysInSheet[k]) {
        sheet.getRange(keysInSheet[k], 2).setValue(settingsMap[k]);
      } else {
        sheet.appendRow([k, settingsMap[k], "마을 설정 요소"]);
      }
    }
    
    // [v48.0] 마을 설정 및 제주시 기상 정보 캐시 강제 무효화 처리 (즉시 기후/BGM 매핑 갱신 보장!)
    try {
      var cache = CacheService.getScriptCache();
      cache.remove("global_village_settings");
      cache.remove("jeju_realtime_weather");
    } catch(err) {}
    
    return { success: true, message: "마을의 기후와 음악 환경이 신비롭게 변화했습니다! 🌌🌦️" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}



/**
 * ==========================================
 * 📚 지혜의 보물고 반응 및 댓글 API
 * ==========================================
 */
function getWisdomTipsWithReactions() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("지혜의_보물고") || ss.insertSheet("지혜의_보물고");
    if (sheet.getLastRow() < 1) {
      sheet.appendRow(["날짜", "제목", "내용", "카테고리", "작성자", "조회수", "공감수", "깨달음수", "사진주소", "기본게시물"]);
      sheet.setFrozenRows(1);
      return [];
    }
    
    var data = sheet.getDataRange().getDisplayValues();
    var commentSheet = ss.getSheetByName("지혜의_보물고_댓글") || ss.insertSheet("지혜의_보물고_댓글");
    var cData = commentSheet.getLastRow() > 1 ? commentSheet.getDataRange().getDisplayValues() : [];
    
    var pKeyRes = getPollinationsApiKey();
    var pKey = pKeyRes.success ? pKeyRes.key : "";
    
    var tips = [];
    for (var i = 1; i < data.length; i++) {
      var tipId = i; // 행 번호 기반 ID
      var comments = [];
      for (var j = 1; j < cData.length; j++) {
        if (Number(cData[j][1]) === tipId) {
          comments.push({
            date: cData[j][0],
            author: cData[j][2],
            content: cData[j][3]
          });
        }
      }
      
      var jVal = data[i].length > 9 ? String(data[i][9] || "").trim() : "";
      var isDefault = (jVal === "Y" || jVal === "true");
      var targetQuest = (jVal !== "Y" && jVal !== "true" && jVal !== "N" && jVal !== "") ? jVal : "";
      
      var rawImage = data[i].length > 8 ? data[i][8] : "";
      var finalImage = appendPollinationsApiKey(rawImage, pKey);
      
      tips.push({
        id: tipId,
        date: data[i][0],
        title: data[i][1],
        content: data[i][2],
        category: data[i][3],
        author: data[i][4] || "길드마스터",
        views: Number(data[i][5] || 0),
        likes: Number(data[i][6] || 0),
        insights: Number(data[i][7] || 0),
        image: finalImage,
        isDefault: isDefault,
        targetQuest: targetQuest,
        comments: comments
      });
    }
    return tips.reverse();
  } catch (e) {
    return [];
  }
}



function getVillageNotices() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get("v58_village_notices");
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      cache.remove("v58_village_notices");
    }
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndInitNoticeSheet(ss);
    
    if (sheet.getLastRow() < 2) return [];
    var data = sheet.getDataRange().getDisplayValues();
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var activeVal = String(data[i][4]).toUpperCase().trim();
      // 아카이브(notice.html)에서는 TRUE/FALSE 상관없이 모든 공지글 내역을 역사처럼 보여줍니다.
      list.push({
        date: data[i][0],      // 작성일
        category: data[i][1],  // 카테고리
        title: data[i][2],     // 제목
        content: data[i][3],   // 내용
        active: (activeVal === "TRUE" || activeVal === "ACTIVATED" || activeVal === "")
      });
    }
    var result = list.reverse(); // 최신 공지 순
    
    // 30분(1800초) 캐시 저장
    cache.put("v58_village_notices", JSON.stringify(result), 1800);
    return result;
  } catch (e) {
    return [];
  }
}




// [이동 완료] SeededRandom 함수는 Common_Utils.gs로 공통 유틸리티 이관되었습니다.




/**
 * ⚡ [v66.0] 하이브리드 돌발 퀘스트 자동화 코어 계산기
 * 날짜(yyyy-MM-dd)를 기반으로 요일, 주말, 임시휴무/공휴일, 월초 결산을 실시간 연산하여 당일 퀘스트를 동적 리턴합니다.
 */
function getAlgorithmicSuddenQuest(dateStr) {
  try {
    var parts = dateStr.split("-");
    var year = parseInt(parts[0], 10);
    var month = parseInt(parts[1], 10);
    var day = parseInt(parts[2], 10);
    
    // 한국 시간(KST) 기준으로 정확하게 날짜 객체 생성
    var targetDate = new Date(year, month - 1, day, 12, 0, 0); 
    
    // 1. 21대 명품 웰니스 테라피 정의 (평일용 - 비타민/미네랄 겹침 방지 최적 순서)
    var baseQuests = [
      { key: "비타민", title: "비타민과 미네랄 day!", description: "활력을 주고 몸의 대사를 돕는 신선한 제철 과일, 채소 섭취나 종합 비타민/영양제 섭취 모습을 인증해주세요! 🍊💊", method: "사진", score: 15 },
      { key: "유산균", title: "건강한 장내환경 day!", description: "유익균이 좋아하는 발효식품(요거트, 청국장, 낫또, 김치)이나 프로바이오틱스 유산균 섭취를 인증해주세요! 🥛🦠", method: "사진", score: 15 },
      { key: "식이섬유", title: "식이섬유 day!", description: "장 건강과 포만감을 책임지는 신선한 쌈채소, 브로콜리, 버섯 등 식이섬유가 풍부한 식단을 드시고 인증해주세요! 🥦🥬", method: "사진", score: 15 },
      { key: "칼슘", title: "칼슘 day!", description: "뼈 건강을 튼튼하게 지켜주는 멸치, 유제품, 두부, 브로콜리 등 칼슘이 풍부한 음식이나 칼슘 영양제 섭취를 인증해 주세요! 🧀🦴", method: "사진", score: 15 },
      { key: "온수", title: "체온면역 day!", description: "체온 1도가 면역력을 좌우합니다! 몸을 따뜻하게 데워줄 순수한 온수/허브티 한 잔이나 족욕, 따뜻한 목욕을 실천하고 인증해 주세요! 🍵♨️", method: "사진", score: 15 },
      { key: "단백질", title: "단백질 day!", description: "근육과 대사를 지켜주는 양질의 단백질(닭가슴살, 흰살생선, 계란, 두부, 살코기 등) 식단을 맛있게 챙겨 드시고 인증해주세요! 🍗🍳", method: "사진", score: 15 },
      { key: "비타민B군", title: "비타민 B군 day!", description: "피로 해소와 에너지 생성을 돕는 비타민 B가 풍부한 식품(돼지고기, 달걀, 시금치 등) 또는 비타민 B 영양제 복용을 인증해주세요! 🥚🔋", method: "사진", score: 15 },
      { key: "간식", title: "건강한 간식 day!", description: "가공간식, 액상과당 음료 대신 견과류, 무가당 요거트, 방울토마토 등 건강한 간식을 선택해 드시고 인증해주세요! 🍅🥜", method: "사진", score: 15 },
      { key: "인슐린", title: "인슐린 day!", description: "식사 시 혈당 폭발을 막기 위해 식이섬유(채소) ➔ 단백질 ➔ 탄수화물 순서로 식사법을 직접 실천하고 인증해 주세요! 🥗📈", method: "사진", score: 15 },
      { key: "영양제", title: "나의 웰니스 보물창고 정리 day!", description: "내가 매일 챙겨 먹는 영양제 보관 상자나 서랍을 깨끗이 정돈하고, 올바른 복용법에 맞춰 가지런히 정돈된 영양제들을 인증해 주세요! 💊📦", method: "사진", score: 15 },
      { key: "프리바이오틱스", title: "프리바이오틱스 day!", description: "장내 유익균의 좋은 먹이가 되는 마늘, 양파, 바나나, 아스파라거스 등 프리바이오틱스 식품이나 영양제 섭취를 인증해 주세요! 🍌🦠", method: "사진", score: 15 },
      { key: "디톡스", title: "디지털 디톡스 day!", description: "오늘 저녁 식사 후 단 1시간 동안 스마트폰을 멀리하고 뇌에 휴식을 주세요! 멀리 치워둔 폰의 모습이나, 폰 대신 읽은 책/가벼운 스트레칭 모습을 인증해 주세요! 📵🧘", method: "사진", score: 15 },
      { key: "미네랄", title: "미네랄 day!", description: "세포 기능을 활성화하는 미네랄이 풍부한 해조류(미역, 다시마), 버섯, 견과류 섭취 모습을 인증해 주세요! 🍄🥜", method: "사진", score: 15 },
      { key: "지방", title: "건강한 지방 day!", description: "불포화지방산이 풍부하여 혈관을 깨끗하게 해주는 아보카도, 견과류, 등푸른 생선, 올리브유 등을 섭취하고 인증해주세요! 🥑🐟", method: "사진", score: 15 },
      { key: "나트륨", title: "나트륨 줄이기 day!", description: "건강한 혈압과 부종 예방을 위해 소금기 적은 삼삼하고 담백한 식단(저염식)을 직접 섭취하고 인증해주세요! 🥗🧂", method: "사진", score: 15 },
      { key: "칼륨", title: "칼륨 day!", description: "나트륨 배출과 부종 완화를 돕는 칼륨이 풍부한 식품(바나나, 토마토, 아보카도, 오이, 시금치 등) 섭취를 인증해주세요! 🥑🥒", method: "사진", score: 15 },
      { key: "항산화", title: "항산화 day!", description: "활성산소를 제거하는 항산화 영양소가 가득한 블루베리, 토마토, 파프리카 등 컬러풀한 베리/야채 섭취를 인증해 주세요! 🫐🍅", method: "사진", score: 15 },
      { key: "수면", title: "건강한 수면 day!", description: "최고의 면역 충전! 침실 어둡게 하기, 온도 조절, 잠들기 전 스마트폰 멀리하기 등 질 좋은 수면을 위한 잠자리 환경을 인증해 주세요! 🛌🌙", method: "사진", score: 15 },
      { key: "생애주기별", title: "생애주기별 필수 영양소 day!", description: "나의 연령/성별 등 생애주기에 꼭 필요하여 특별히 챙겨 드시는 영양 식품이나 영양제를 인증해주세요! 🍊🍼", method: "사진", score: 15 },
      { key: "탄수화물", title: "착한 탄수화물 day!", description: "정제되지 않아 식이섬유가 살아있는 현미밥, 고구마, 단호박, 통밀 등 착한 탄수화물 섭취를 인증해 주세요! 🍠🌾", method: "사진", score: 15 },
      { key: "지중해", title: "지중해 식단 day!", description: "세계가 인정한 장수 식단! 신선한 야채, 올리브유, 견과류, 통곡물이 어우러진 건강한 지중해 스타일의 식사를 즐기고 인증해 주세요! 🥗🍋", method: "사진", score: 15 }
    ];
    
    // 2. 월별 시드값을 기반으로 순서는 유지한 채 시작 위치(Offset)만 변경 (상대적 간격 고정)
    var offset = (year * 12 + month) % 21;
    var shuffledQuests = [];
    for (var i = 0; i < 21; i++) {
      shuffledQuests.push(baseQuests[(i + offset) % 21]);
    }
    
    // 3. 당월 1일부터 주어진 targetDate까지 하루씩 전진하며 배정 알고리즘 구동
    var rotationIdx = 0;
    var countMap = {};
    var usedDays = []; // 월초 결산/목표세우기용 사용 가능 평일 추적
    
    // 해당 월의 1일부터 targetDate까지 매일의 퀘스트를 계산
    for (var d = 1; d <= day; d++) {
      var curDate = new Date(year, month - 1, d, 12, 0, 0);
      var curDayOfWeek = curDate.getDay(); // 0:일, 1:월, ..., 6:토
      var curDateStr = year + "-" + (month < 10 ? "0" + month : month) + "-" + (d < 10 ? "0" + d : d);
      
      // A. 주말 미션 (토요일 2대 로테이션, 일요일 3대 로테이션)
      if (curDayOfWeek === 6) {
        if (d === day) {
          var satIdx = Math.floor((d - 1) / 7) % 2;
          if (satIdx === 0) {
            return {
              title: "주말 미션! 홈트 & 운동 day!",
              description: "집에서 가벼운 홈트레이닝을 하거나 스포츠(필라테스, 헬스, 등산, 배드민턴 등)를 즐기는 내 운동 인증을 올려 주세요! 🏋️‍♂️💪",
              method: "사진",
              score: 15
            };
          } else {
            return {
              title: "운동 전후 영양관리 day!",
              description: "운동 전 활력을 주는 영양이나 운동 후 회복을 위한 단백질/수분 보완 식사 인증을 올려 주세요! 🏃‍♀️🍼",
              method: "사진",
              score: 15
            };
          }
        }
        continue;
      }
      if (curDayOfWeek === 0) {
        if (d === day) {
          var sunIdx = Math.floor((d - 1) / 7) % 3;
          if (sunIdx === 0) {
            return {
              title: "주말 미션! 초록 힐링 야외 산책 day!",
              description: "일요일은 자연 속에서 뇌를 비우는 날! 공원이나 산책길에서 마주한 초록 풍경 또는 산책하는 내 모습을 인증해 주세요! 🌳🚶‍♀️",
              method: "사진",
              score: 15
            };
          } else if (sunIdx === 1) {
            return {
              title: "올바른 물마시기 day!",
              description: "주말 동안 흐트러진 수분 밸런스를 맞추는 날! 맑은 물 한 잔을 마시는 건강한 모습과 산책길을 인증해 주세요! 🥛👣",
              method: "사진",
              score: 15
            };
          } else {
            return {
              title: "수분균형 day!",
              description: "몸 속 완벽한 수분 밸런스 유지! 수시로 충분히 물을 마시고 텀블러나 물병과 함께 편히 휴식하는 인증을 올려 주세요! 💧🛋️",
              method: "사진",
              score: 15
            };
          }
        }
        continue;
      }
      
      // B. 임시 지정 휴무일 및 법정 공휴일 (주말 미션 우회 배정)
      if (isCenterHoliday(curDate)) {
        if (d === day) {
          if (d % 2 === 1) {
            return {
              title: "주말 미션! 홈트 & 운동 day!",
              description: "센터 휴무일이지만 우리의 웰니스는 계속됩니다! 가벼운 홈트레이닝이나 야외 스포츠를 즐기는 건강한 에너지를 멋지게 인증해 주세요! 🏃‍♂️💪",
              method: "사진",
              score: 15
            };
          } else {
            return {
              title: "주말 미션! 초록 힐링 야외 산책 day!",
              description: "센터 휴무일을 맞아 마음에 힐링을 주는 시간! 근처 공원이나 산책길을 걸으며 마주한 아름다운 자연 풍경이나 싱그러운 산책길 모습을 인증해 주세요! 🌳🚶‍♀️",
              method: "사진",
              score: 15
            };
          }
        }
        continue;
      }
      
      // C. 주중 사용 가능한 평일 체크 (목요일 포함!)
      usedDays.push(d);
      var isMonthlyReviewDay = (usedDays.length === 1);
      var isGoalSettingDay = (usedDays.length === 2);
      
      // 🚀 정식 출시 2026년 6월 특별 예외 필터링 (5월 결산 생략 ➡️ 1일에 즉시 목표 세우기 가동!)
      var isLaunchMonth = (year === 2026 && month === 6);
      if (isLaunchMonth) {
        isMonthlyReviewDay = false; // 5월 결산 완벽 스킵
        isGoalSettingDay = (usedDays.length === 1); // 6월 1일에 즉시 목표 세우기 매핑!
      }
      
      var prevMonth = month === 1 ? 12 : month - 1;
      
      // E. 단일 매핑 처리 (목요일 주간 결산일은 폐지하고 일반 평일 로테이션 미션으로 대체)
      if (isMonthlyReviewDay) {
        if (d === day) {
          return {
            title: prevMonth + "월 월간 결산일!",
            description: "지난 " + prevMonth + "월 한 달간의 나의 웰니스 점수와 성취도를 최종 점검하고 나 스스로에 대한 솔직한 평가와 응원의 한마디를 오아시스에 남겨보세요! 🏆✨",
            method: "게시판",
            score: 15
          };
        }
        continue;
      }
      
      if (isGoalSettingDay) {
        if (d === day) {
          return {
            title: month + "월 목표 세우기!",
            description: "새로운 " + month + "월이 시작되었습니다! 이번 한 달간 꼭 이루고 싶은 나의 웰니스/체중 감량 목표와 건강 루틴을 세우고 오아시스에 당차게 선포해 보세요! 🎯📝",
            method: "게시판",
            score: 15
          };
        }
        continue;
      }
      
      // E. 일반 평일 21대 테라피 로테이션
      var quest = shuffledQuests[rotationIdx % 21];
      var count = countMap[quest.key] || 0;
      
      // ⚠️ 영양제 정리 미션은 한 달에 최대 1회만 출현하도록 제한하는 최적화 건너뛰기 룰
      if (quest.key === "영양제" && count >= 1) {
        rotationIdx++;
        quest = shuffledQuests[rotationIdx % 21];
        count = countMap[quest.key] || 0;
      }
      
      count++;
      countMap[quest.key] = count;
      
      if (d === day) {
        var finalTitle = (count > 1) ? count + "차 " + quest.title : quest.title;
        return {
          title: finalTitle,
          description: quest.description,
          method: quest.method,
          score: quest.score
        };
      }
      rotationIdx++;
    }
    
    return null;
  } catch (err) {
    Logger.log("getAlgorithmicSuddenQuest 에러: " + err.toString());
    return null;
  }
}



function getActiveQuestStatus(phone, ss, logData, memberName) {
  var result = {
    todayQuest: null,
    tomorrowQuest: null,
    glycogenQuest: null,
    shield: { active: false, expireStr: "" }
  };
  
  try {
    // [perf] ss를 인자로 받지 못했을 경우에만 새로 연결 (독립 호출 호환)
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndCreateQuestRegistrySheet();
    var data = sheet.getDataRange().getDisplayValues();
    
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    var tomorrow = new Date(now.getTime() + 24 * 60 * 60 * 1000);
    var tomorrowStr = Utilities.formatDate(tomorrow, "GMT+9", "yyyy-MM-dd");
    
    var cleanPhone = formatPhoneNumber(phone).replace(/[^0-9]/g, "");
    
    // [perf] 3개의 for-loop을 1개로 통합하여 퀘스트 시트 3중 반복 스캔을 제거
    var activeGlycogenRowIdx = -1;
    var latestGlycogenQuest = null;
    var latestSuccessTime = 0;

    for (var i = 1; i < data.length; i++) {
      var type = data[i][2];   // C열: 유형
      var dateStr = data[i][1]; // B열: 시행일
      var title = data[i][3];  // D열: 퀘스트명
      var desc = data[i][4];   // E열: 설명
      var qPhone = formatPhoneNumber(data[i][6]).replace(/[^0-9]/g, ""); // G열: 전화번호
      var status = data[i][7]; // H열: 상태

      // 1. 이장 공지 퀘스트
      if (type === "이장") {
        var qScore = (data[i] && data[i].length > 8) ? Number(data[i][8] || 15) : 15;
        var qMethod = (data[i] && data[i].length > 9) ? String(data[i][9] || '사진').trim() : '사진';
        
        // ⚠️ 이중 방어막: 제목에 명시적인 이모지가 있다면 백엔드 단에서 강제로 매핑해 주는 인텔리전트 보정
        var titleStr = String(title || "");
        if (titleStr.indexOf('✍️') > -1 || titleStr.indexOf('📝') > -1 || titleStr.indexOf('✏️') > -1 || titleStr.indexOf('⚡') > -1) {
          qMethod = '게시판';
        } else if (titleStr.indexOf('📸') > -1 || titleStr.indexOf('📷') > -1) {
          qMethod = '사진';
        }

        if (dateStr === todayStr) {
          // 💡 [v66.2] 지식창고 연동 넛지 배너 동적 결합 (초슬림 한 줄 넛지)
          var hasRelatedTip = false;
          try {
            var keywordList = ["식이섬유", "수분", "물", "오운완", "단식", "체지방", "근육", "단백질", "유산소", "스트레칭", "습관", "식단", "채소"];
            var detectedKeyword = "";
            for (var k = 0; k < keywordList.length; k++) {
              var kw = keywordList[k];
              if (title.indexOf(kw) > -1) {
                detectedKeyword = kw;
                break;
              }
            }
            
            var normalizeTextForMatch = function(str) {
              if (!str) return "";
              return String(str).toLowerCase().replace(/\s+/g, "").replace(/[!?,.『』'"]/g, "").trim();
            };
            var normalizedTodayQuest = normalizeTextForMatch(title);
            
            if (detectedKeyword || normalizedTodayQuest) {
              var wisdomSheet = ss.getSheetByName("지혜의_보물고");
              if (wisdomSheet && wisdomSheet.getLastRow() > 1) {
                var wData = wisdomSheet.getDataRange().getDisplayValues();
                for (var w = 1; w < wData.length; w++) {
                  var wCategory = String(wData[w][3] || "").trim();
                  var wTitle = String(wData[w][1] || "").trim();
                  var wContent = String(wData[w][2] || "").trim();
                  var jVal = wData[w].length > 9 ? String(wData[w][9] || "").trim() : "";
                  var targetQuest = (jVal !== "Y" && jVal !== "true" && jVal !== "N" && jVal !== "") ? jVal : "";
                  
                  // 0순위 J열 타겟팅 매칭
                  if (targetQuest) {
                    var normalizedTarget = normalizeTextForMatch(targetQuest);
                    if (normalizedTarget && normalizedTodayQuest.indexOf(normalizedTarget) > -1) {
                      hasRelatedTip = true;
                      break;
                    }
                  }
                  // 1순위 키워드 매칭
                  if (detectedKeyword) {
                    if (wTitle.indexOf(detectedKeyword) > -1 || wCategory.indexOf(detectedKeyword) > -1) {
                      hasRelatedTip = true;
                      break;
                    }
                  }
                }
              }
            }
          } catch (e) {
            Logger.log("⚠️ 배너 넛지 매칭 수집 중 오류: " + e.toString());
          }

          var finalDesc = desc;
          if (hasRelatedTip) {
            finalDesc = desc + "\n\n📖 오늘의 퀘스트 관련 지식칼럼도 지식창고에서 확인해보세요 🌿";
          }

          result.todayQuest = { title: title, description: finalDesc, score: qScore, method: qMethod };
        }
      }

      // 2 & 3. 글리코겐/요요방패 (해당 회원 것만)
      if (type === "글리코겐" && qPhone === cleanPhone) {
        var triggerTimeStr = data[i][5]; // F열: 만료일
        var limitTime = new Date(data[i][5]);

        if (status === "진행중") {
          if (now > limitTime) {
            sheet.getRange(i + 1, 8).setValue("실패");
            
            // 😢 글리코겐 방패 실패 위로 쪽지 전송 (1회)
            try {
              var mName = memberName;
              if (!mName) {
                var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
                if (regSheet) {
                  var regData = regSheet.getDataRange().getDisplayValues();
                  var regCols = getRegColumnIndices(regSheet);
                  for (var r = 1; r < regData.length; r++) {
                    var rPhone = formatPhoneNumber(regData[r][regCols.phone]).replace(/[^0-9]/g, "");
                    if (rPhone === cleanPhone) {
                      mName = regData[r][regCols.name];
                      break;
                    }
                  }
                }
              }
              if (!mName) mName = "회원";
              mName = String(mName).replace(/\d{4}$/, ""); // 이름 뒤 숫자 제거
              
              if (!hasSentNotificationToday(cleanPhone, "방어", "방패 실패")) { sendPersonalNotification(
                cleanPhone,
                "방어",
                "😢 아쉽습니다! 글리코겐 방패 충전 실패",
                mName + "님, 아쉽게도 3일 제한 시간 내에 미션을 완수하지 못해 글리코겐 방패 충전에 실패하셨습니다. 😢\n\n비록 이번 방패는 충전하지 못했지만 낙담하지 마세요! 건강한 습관을 다시 채우는 것이 중요합니다. 오늘부터 다시 가볍고 건강한 식단과 함께 클럽에서 땀 흘려보아요. 언제나 회원님의 도전을 응원합니다! ❤️"
              );
              } } catch (errFailNoti) {
              console.error("Error sending shield fail notification: " + errFailNoti.toString());
            }
          } else {
            activeGlycogenRowIdx = i + 1;
            latestGlycogenQuest = {
              rowIdx: i + 1,
              startStr: dateStr,
              deadlineStr: triggerTimeStr,
              deadlineTime: limitTime.getTime()
            };
          }
        } else if (status === "성공") {
          if (now <= limitTime && limitTime.getTime() > latestSuccessTime) {
            latestSuccessTime = limitTime.getTime();
          }
        }
      }
    }

    if (latestSuccessTime > 0) {
      result.shield.active = true;
      result.shield.expireStr = Utilities.formatDate(new Date(latestSuccessTime), "GMT+9", "yyyy-MM-dd HH:mm:ss");
    }
    
    // 4. If Glycogen Quest is active, calculate its progress (number of unique attendances since starting)
    if (latestGlycogenQuest) {
      var attendDates = new Set();
      var triggerDateOnlyStr = latestGlycogenQuest.startStr.split(" ")[0]; // e.g. "2026-05-18"
      
      // logData가 없으면 새로 로드 (독립 호출 호환용)
      var localLogData = logData;
      if (!localLogData) {
        var logSheet = ss.getSheetByName("출석기록");
        if (logSheet) {
          localLogData = logSheet.getDataRange().getValues();
        }
      }
      
      if (localLogData) {
        for (var j = 1; j < localLogData.length; j++) {
          var logPhone = formatPhoneNumber(localLogData[j][3]).replace(/[^0-9]/g, ""); // D열: 전화번호
          if (logPhone === cleanPhone) {
            var logDateRaw = localLogData[j][0]; // A열: 날짜
            var logDateStr = (logDateRaw instanceof Date) ? Utilities.formatDate(logDateRaw, "GMT+9", "yyyy-MM-dd") : String(logDateRaw);
            
            // Count unique dates starting from the trigger date (next day or same day)
            if (logDateStr >= triggerDateOnlyStr) {
              attendDates.add(logDateStr);
            }
          }
        }
      }
      
      var attendCount = attendDates.size;
      if (attendCount >= 3) {
        // SUCCESS!
        sheet.getRange(latestGlycogenQuest.rowIdx, 8).setValue("성공");
        var shieldExpire = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);
        sheet.getRange(latestGlycogenQuest.rowIdx, 6).setValue(Utilities.formatDate(shieldExpire, "GMT+9", "yyyy-MM-dd HH:mm:ss"));
        
        // memberName이 없으면 등록현황에서 로드 (독립 호출 호환용)
        var mName = memberName;
        if (!mName) {
          var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
          if (regSheet) {
            var regData = regSheet.getDataRange().getDisplayValues();
            var regCols = getRegColumnIndices(regSheet);
            for (var r = 1; r < regData.length; r++) {
              var rPhone = String(regData[r][regCols.phone]).replace(/[^0-9]/g, "");
              if (rPhone === cleanPhone) {
                mName = regData[r][regCols.name];
                break;
              }
            }
          }
        }
        if (!mName) mName = "모험가";
        
        recordActivityLog({
          phone: cleanPhone,
          name: mName,
          type: "방어",
          item: "🔥 글리코겐 클리어 완료",
          score: 100,
          action: "성공",
          statType: "def"
        });
        
        var archiveSheet = ss.getSheetByName("아카이브");
        if (archiveSheet) {
          var dateStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
          var timeStr = Utilities.formatDate(now, "GMT+9", "HH:mm:ss");
          var cardComment = `🎉 [글리코겐 클리어 성공] ${mName}님이 3일 연속 지옥의 출석으로 치팅데이를 완전히 격파하고 '요요 방어 방패'를 획득하셨습니다! 🛡️ (+100 EXP)`;
          archiveSheet.appendRow([
            dateStr,
            timeStr,
            mName,
            "'" + cleanPhone,
            "퀘스트",
            "🔥 글리코겐 클리어 성공",
            cardComment,
            "SYSTEM_CARD",
            100
          ]);
        }
        
        // 🛡️ 3일차 성공 쪽지 발송 (중복 방지)
        try {
          var cleanPhoneName = mName || "회원";
          cleanPhoneName = String(cleanPhoneName).replace(/\d{4}$/, ""); // 이름 뒤 숫자 제거
          if (!hasSentNotificationToday(cleanPhone, "방어", "방패 완성")) {
            sendPersonalNotification(
              cleanPhone,
              "방어",
              "🛡️ 요요 방어 방패 완성! 🎉",
              cleanPhoneName + "님, 축하합니다! 3일 연속 지옥의 미션 출석을 완수하여 치팅데이를 완전히 극복하고 '요요 방어 방패'를 완성하셨습니다! 🛡️\n\n앞으로 7일 동안 요요와 글리코겐의 공격으로부터 완벽하게 보호됩니다! 이 기세를 몰아 건강한 습관을 계속 유지해보아요! 화이팅! ❤️"
            );
          }
        } catch (errNoti) {
          console.error("Error sending shield success notification: " + errNoti.toString());
        }

        result.glycogenQuest = null;
        result.shield.active = true;
        result.shield.expireStr = Utilities.formatDate(shieldExpire, "GMT+9", "yyyy-MM-dd HH:mm:ss");
      } else {
        // 🛡️ 1, 2일차 진행도 쪽지 발송 (중복 방지)
        try {
          var mName = memberName;
          if (!mName) {
            var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
            if (regSheet) {
              var regData = regSheet.getDataRange().getDisplayValues();
              var regCols = getRegColumnIndices(regSheet);
              for (var r = 1; r < regData.length; r++) {
                var rPhone = String(regData[r][regCols.phone]).replace(/[^0-9]/g, "");
                if (rPhone === cleanPhone) {
                  mName = regData[r][regCols.name];
                  break;
                }
              }
            }
          }
          if (!mName) mName = "회원";
          mName = String(mName).replace(/\d{4}$/, ""); // 이름 뒤 숫자 제거

          if (attendCount === 1) {
            if (!hasSentNotificationToday(cleanPhone, "방어", "1일차 완성")) {
              sendPersonalNotification(
                cleanPhone,
                "방어",
                "🛡️ 글리코겐 방패 1일차 조각 완성!",
                mName + "님, 치팅 극복을 위한 위대한 첫 걸음을 내딛으셨습니다! 🏃‍♀️\n\n1일차 글리코겐 방패 조각이 완성되었습니다. 🛡️ 치팅 데이의 흔적을 깨끗이 날려버릴 수 있도록 내일도 꼭 출석하셔서 2일차 조각을 완성해보세요! 클럽에서 기다리고 있겠습니다! 화이팅! ❤️"
              );
            }
          } else if (attendCount === 2) {
            if (!hasSentNotificationToday(cleanPhone, "방어", "2일차 완성")) {
              sendPersonalNotification(
                cleanPhone,
                "방어",
                "🛡️ 글리코겐 방패 2일차 조각 완성!",
                "와우, 대단하십니다, " + mName + "님! 😍\n\n벌써 2일차 방패 조각까지 완성되었습니다! 🛡️ 이제 단 하루만 더 출석하시면 치팅을 완벽하게 진압하고 요요로부터 지켜줄 강력한 '요요 방어 방패'가 완성됩니다. 마지막 조각을 완성하러 내일 꼭 클럽에서 만나요! 🔥"
              );
            }
          }
        } catch (errProgressNoti) {
          console.error("Error sending shield progress notification: " + errProgressNoti.toString());
        }

        result.glycogenQuest = {
          progress: attendCount,
          deadlineStr: latestGlycogenQuest.deadlineStr,
          startStr: latestGlycogenQuest.startStr
        };
      }
    }
    // ⚠️ [v66.0] 오늘의 돌발 퀘스트 누락 시 안전한 동적 백필 (자가치유 방어막)
    if (!result.todayQuest) {
      var algoQuest = getAlgorithmicSuddenQuest(todayStr);
      if (algoQuest) {
        result.todayQuest = {
          title: algoQuest.title,
          description: algoQuest.description,
          score: algoQuest.score || 15,
          method: algoQuest.method || "사진"
        };
      }
    }
  } catch (e) {
    console.error("Error in getActiveQuestStatus:", e.toString());
  }
  
  return result;
}


// ──────────────────────────────────────────────
// [이관] getScheduledQuests
// ──────────────────────────────────────────────
function getScheduledQuests() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndCreateQuestRegistrySheet();
    var data = sheet.getDataRange().getDisplayValues();
    
    // 1. 스프레드시트에 기 적재되어 있는 모든 수동 예약 퀘스트들을 맵(Map)에 적재 (날짜 표준화 적용)
    var manualMap = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][2] === "이장") {
        var rawDate = data[i][1];
        var dStr = normalizeDateStr(rawDate);
        if (dStr) {
          manualMap[dStr] = {
            rowIdx: i + 1,
            date: dStr, // 표준 포맷 전달
            title: data[i][3],
            description: data[i][4],
            status: data[i][7] || "대기",
            score: data[i].length > 8 ? Number(data[i][8] || 15) : 15,
            method: data[i].length > 9 ? (data[i][9] || "사진") : "사진",
            isManual: true
          };
        }
      }
    }
    
    // 2. 오늘부터 앞으로 45일간의 날짜에 대해 수동 또는 자동화 퀘스트 매칭
    var list = [];
    var now = new Date();
    // 타임존 오프셋 반영하여 KST 기준 오늘 날짜 구하기 (더블 오프셋 버그 방지를 위해 포맷 후 재파싱)
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var parts = todayStr.split("-");
    var todayKst = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 12, 0, 0);
    
    for (var d = 0; d < 45; d++) {
      var targetDate = new Date(todayKst.getTime() + d * 24 * 60 * 60 * 1000);
      var targetDateStr = Utilities.formatDate(targetDate, "GMT+9", "yyyy-MM-dd");
      
      if (manualMap[targetDateStr]) {
        // 수동 예약이 이미 존재하는 경우
        list.push(manualMap[targetDateStr]);
      } else {
        // 수동 예약이 없는 미래 날짜 ➡️ 실시간 계산기를 돌려 자동화 가상 예약 주입!
        var algoQuest = getAlgorithmicSuddenQuest(targetDateStr);
        if (algoQuest) {
          list.push({
            rowIdx: null,
            date: targetDateStr,
            title: algoQuest.title,
            description: algoQuest.description,
            status: "진행",
            score: algoQuest.score || 15,
            method: algoQuest.method || "사진",
            isManual: false
          });
        }
      }
    }
    
    // 시간 오름차순 (가까운 날짜 순) 정렬하여 반환
    return list;
  } catch (e) {
    Logger.log("getScheduledQuests 에러: " + e.toString());
    return [];
  }
}

// ──────────────────────────────────────────────
// [이관] saveScheduledQuest
// ──────────────────────────────────────────────
function saveScheduledQuest(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndCreateQuestRegistrySheet();
    
    var dateStr = normalizeDateStr(payload.date); // 확실하게 yyyy-MM-dd 표준형으로 정리
    var title = payload.title;
    var desc = payload.description;
    var score = Number(payload.score || 15);
    var method = payload.method || "사진";
    
    // 동일한 날짜에 이미 등록된 이장 돌발 퀘스트가 있는지 검사하여 덮어쓰기/수정 지원 (날짜 정합성 매핑)
    var data = sheet.getDataRange().getDisplayValues();
    var foundRowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      var sheetDateStr = normalizeDateStr(data[i][1]);
      if (data[i][2] === "이장" && sheetDateStr === dateStr) {
        foundRowIdx = i + 1;
        break;
      }
    }
    
    if (foundRowIdx > -1) {
      // 덮어쓰기 수정 (날짜도 표준 포맷인 dateStr로 예쁘게 치환)
      sheet.getRange(foundRowIdx, 2).setValue(dateStr);
      sheet.getRange(foundRowIdx, 4).setValue(title);
      sheet.getRange(foundRowIdx, 5).setValue(desc);
      sheet.getRange(foundRowIdx, 8).setValue("대기");
      sheet.getRange(foundRowIdx, 9).setValue(score);
      sheet.getRange(foundRowIdx, 10).setValue(method);
    } else {
      // 신규 등록
      var newId = "Q_" + Date.now();
      sheet.appendRow([
        newId,
        dateStr,
        "이장",
        title,
        desc,
        dateStr + " 23:59:59", // 만료 시간
        "ALL_MEMBERS",
        "대기",
        score,
        method
      ]);
    }
    
    return { success: true, message: dateStr + "일 당일에 자동 선포되는 돌발 퀘스트 예약 등록이 안전하게 완료되었습니다! 📅⚡" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ──────────────────────────────────────────────
// [이관] deleteScheduledQuest
// ──────────────────────────────────────────────
function deleteScheduledQuest(rowIdx) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndCreateQuestRegistrySheet();
    sheet.deleteRow(Number(rowIdx));
    return { success: true, message: "예약된 돌발 퀘스트가 캘린더에서 제거되었습니다." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ──────────────────────────────────────────────
// [이관] processScheduledQuestsDaily
// ──────────────────────────────────────────────
function processScheduledQuestsDaily() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndCreateQuestRegistrySheet();
    var data = sheet.getDataRange().getDisplayValues();
    
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var activatedCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      var rawDate = data[i][1];
      var dateStr = normalizeDateStr(rawDate); // 날짜 포맷 강제 보정
      var type = data[i][2];    // C열: 유형 ("이장")
      var title = data[i][3];   // D열: 퀘스트명
      var desc = data[i][4];    // E열: 설명
      var status = data[i][7];  // H열: 상태 ("대기")
      var score = Number(data[i][8] || 15);
      var method = data[i][9] || "사진";
      
      if (type === "이장" && dateStr === todayStr && status === "대기") {
        var rowIdx = i + 1;
        // 1. 날짜 데이터가 비표준형인 경우 표준형으로 일괄 등기 보정
        sheet.getRange(rowIdx, 2).setValue(todayStr);
        // 2. 상태를 '진행'으로 변경하여 당일 퀘스트로 공식 표출
        sheet.getRange(rowIdx, 8).setValue("진행");
        
        // 3. 당일 아침 활성화 알림(쪽지)을 회원 전체에게 1회 공식 선포!
        var globalTitle = "📅 [오늘의 돌발 퀘스트] " + title + " ⚡";
        var globalContent = "📢 [웰니스 코치의 돌발 퀘스트 오늘 진행 시작!] 📢\n\n" +
                            "오늘의 돌발 퀘스트가 공식적으로 활성화되었습니다! 오늘의 웰니스 수호에 도전해 보세요!\n\n" +
                            "📅 [시행 일자]: " + todayStr + " (오늘 밤 자정 마감)\n" +
                            "🔥 [돌발 퀘스트]: " + title + "\n" +
                            "📝 [임무 설명]: " + desc + "\n" +
                            "💎 [보상 EXP]: +" + score + " EXP\n" +
                            "🎯 [인증 방법]: " + method + " 인증\n\n" +
                            "오늘 밤 자정 전까지 아카이브 게시판에 인증을 완료해 주세요. 수호 점수 획득을 향해 돌격! ⚔️";
                            
        sendGlobalNotification("quest", globalTitle, globalContent);
        activatedCount++;
        Logger.log("✅ [돌발퀘스트 활성화] 오늘 자(" + todayStr + ") 퀘스트 '" + title + "' 공식 진행 개시 및 알림 발송 완료!");
      }
    }
    
    return { success: true, activatedCount: activatedCount };
  } catch (e) {
    Logger.log("❌ processScheduledQuestsDaily 에러: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ──────────────────────────────────────────────
// [이관] getDailyQuestNoticePreview
// ──────────────────────────────────────────────
function getDailyQuestNoticePreview() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("돌발퀘스트_목록");
    if (!sheet) {
      return { success: false, error: "돌발퀘스트_목록 시트가 없습니다." };
    }
    var data = sheet.getDataRange().getDisplayValues();
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    var todayQuest = null;
    var isManual = false;
    
    for (var i = 1; i < data.length; i++) {
      var type = String(data[i][2] || "").trim();
      var rawDate = data[i][1];
      var dateStr = normalizeDateStr(rawDate); // 날짜 정합성 표준화
      if (type === "이장" && dateStr === todayStr) {
        var title = String(data[i][3] || "").trim();
        var desc = String(data[i][4] || "").trim();
        var qScore = (data[i] && data[i].length > 8) ? Number(data[i][8] || 15) : 15;
        var qMethod = (data[i] && data[i].length > 9) ? String(data[i][9] || '사진').trim() : '사진';
        
        // ⚠️ 이중 방어막: 제목에 명시적인 이모지가 있다면 백엔드 단에서 강제로 매핑해 주는 인텔리전트 보정
        var titleStr = String(title || "");
        if (titleStr.indexOf('✍️') > -1 || titleStr.indexOf('📝') > -1 || titleStr.indexOf('✏️') > -1 || titleStr.indexOf('⚡') > -1) {
          qMethod = '게시판';
        } else if (titleStr.indexOf('📸') > -1 || titleStr.indexOf('📷') > -1) {
          qMethod = '사진';
        }
        
        todayQuest = { title: title, description: desc, score: qScore, method: qMethod };
        isManual = true;
        break;
      }
    }
    
    // 만약 오늘자 수동 예약이 없다면 ➡️ 알고리즘 계산기 작동하여 가상 프리뷰 생성!
    if (!todayQuest) {
      var algoQuest = getAlgorithmicSuddenQuest(todayStr);
      if (algoQuest) {
        todayQuest = {
          title: algoQuest.title,
          description: algoQuest.description,
          score: algoQuest.score || 15,
          method: algoQuest.method || "사진"
        };
      }
    }
    
    if (!todayQuest) {
      return { success: false, error: "오늘(" + todayStr + ") 자 돌발 퀘스트를 판정할 수 없습니다." };
    }
    
    // 💡 [v66.2] 원장님의 실천 철학: 돌발 퀘스트와 관련된 선제작 지식창고 글이 있을 때만 동적으로 특별 비책 가이드를 우편 꼬리에 넛지로 동봉!
    // (매칭 칼럼이 전혀 없는 날에는 완전히 숨김 처리하여 메시지 초슬림화 사수)
    var tipNotice = "";
    try {
      // A. 돌발퀘스트 제목에서 다이어트 주요 키워드 추출
      var keywordList = ["식이섬유", "수분", "물", "오운완", "단식", "체지방", "근육", "단백질", "유산소", "스트레칭", "습관", "식단", "채소"];
      var detectedKeyword = "";
      for (var k = 0; k < keywordList.length; k++) {
        var kw = keywordList[k];
        if (todayQuest.title.indexOf(kw) > -1) {
          detectedKeyword = kw;
          break;
        }
      }
      
      var normalizeTextForMatch = function(str) {
        if (!str) return "";
        return String(str).toLowerCase().replace(/\s+/g, "").replace(/[!?,.『』'"]/g, "").trim();
      };
      var normalizedTodayQuest = normalizeTextForMatch(todayQuest.title);
      
      if (detectedKeyword || normalizedTodayQuest) {
        // B. 지혜의_보물고 시트에서 관련 글 매칭 조회
        var wisdomSheet = ss.getSheetByName("지혜의_보물고");
        if (wisdomSheet && wisdomSheet.getLastRow() > 1) {
          var wData = wisdomSheet.getDataRange().getDisplayValues();
          var relatedTitles = [];
          
          for (var w = 1; w < wData.length; w++) {
            var wCategory = String(wData[w][3] || "").trim();
            var wTitle = String(wData[w][1] || "").trim();
            var jVal = wData[w].length > 9 ? String(wData[w][9] || "").trim() : "";
            var targetQuest = (jVal !== "Y" && jVal !== "true" && jVal !== "N" && jVal !== "") ? jVal : "";
            
            var isMatch = false;
            
            // 0순위 J열 타겟팅 매칭
            if (targetQuest) {
              var normalizedTarget = normalizeTextForMatch(targetQuest);
              if (normalizedTarget && normalizedTodayQuest.indexOf(normalizedTarget) > -1) {
                isMatch = true;
              }
            }
            
            // 1순위 키워드 매칭 (제목/카테고리만 매칭, 본문 내용 검색은 제외)
            if (!isMatch && detectedKeyword) {
              if (wTitle.indexOf(detectedKeyword) > -1 || wCategory.indexOf(detectedKeyword) > -1) {
                isMatch = true;
              }
            }
            
            // 섹션 1 고정글 제외 및 매칭 칼럼 수집
            if (isMatch && wCategory !== "섹션 1") {
              relatedTitles.push(wTitle);
            }
          }
          
          // C. 매칭 칼럼이 존재할 때만 원장님 조율 피드백을 수용한 단정한 넛지 단락 합성!
          if (relatedTitles.length > 0) {
            tipNotice = "\n\n오늘 돌발퀘스트와 관련한 지식 칼럼들을 지식창고에서 확인할수 있습니다.\n" +
                        "추천칼럼: " + relatedTitles.join(", ") + " 등 지식창고 탭을 통해 관련 건강정보를 확인하세요";
          }
        }
      }
    } catch (e) {
      Logger.log("⚠️ 쪽지 발송 지식창고 연동 꿀팁 추출 에러: " + e.toString());
    }

    var title = "⚡ [오늘의 돌발 퀘스트] " + todayQuest.title + " ✉️";
    var content = "📢 [이장님의 긴급 돌발 퀘스트 선포!] 📢\n\n" +
                  "오늘의 돌발 퀘스트가 공식적으로 선포되었습니다! 오늘 자정 전까지 완수하고 추가 경험치를 획득해 보세요!\n\n" +
                  "📅 [시행 일자]: " + todayStr + " (오늘 밤 자정 마감)\n" +
                  "🔥 [돌발 퀘스트]: " + todayQuest.title + "\n" +
                  "📝 [임무 설명]: " + todayQuest.description + "\n" +
                  "💎 [보상 EXP]: +" + todayQuest.score + " EXP\n" +
                  "🎯 [인증 방법]: " + todayQuest.method + " 인증" +
                  tipNotice + "\n\n" +
                  "지금 즉시 대시보드에서 아코디언 서랍을 열어 퀘스트 상세 가이드를 확인하고 도전해 보세요! ⚔️";
                  
    return { success: true, title: title, content: content, quest: todayQuest, isManual: isManual };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ──────────────────────────────────────────────
// [이관] autoSendDailyQuestNotice
// ──────────────────────────────────────────────
function autoSendDailyQuestNotice() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndCreateQuestRegistrySheet();
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    // 1. 프리뷰 엔진 작동 (수동이 없으면 가상 알고리즘 퀘스트 탑재)
    var preview = getDailyQuestNoticePreview();
    if (!preview || !preview.success) {
      Logger.log("돌발 퀘스트 프리뷰 생성 실패로 트리거 종료: " + (preview ? preview.error : ""));
      return;
    }
    
    var questInfo = preview.quest;
    var isManual = preview.isManual;
    
    // 2. 수동과 자동 분기 처리하여 시트 등기/진행 상태 기록
    if (isManual) {
      // 오늘자 수동 예약을 찾아 '진행' 상태로 변경 및 표준화 저장
      var data = sheet.getDataRange().getDisplayValues();
      for (var i = 1; i < data.length; i++) {
        var sheetDateStr = normalizeDateStr(data[i][1]);
        if (data[i][2] === "이장" && sheetDateStr === todayStr) {
          sheet.getRange(i + 1, 2).setValue(todayStr); // 표준 yyyy-MM-dd 포맷 보정 저장
          sheet.getRange(i + 1, 8).setValue("진행"); // H열: 상태를 진행으로 세팅
          Logger.log("✅ 오늘 자 수동 예약 돌발 퀘스트를 '진행'으로 공식 활성화 완료!");
          break;
        }
      }
    } else {
      // 오늘자 자동 퀘스트를 시트(돌발퀘스트_목록)에 등기 적재 (appendRow)
      var newId = "Q_AUTO_" + Date.now();
      sheet.appendRow([
        newId,
        todayStr,
        "이장",
        questInfo.title,
        questInfo.description,
        todayStr + " 23:59:59", // 만료 시간
        "ALL_MEMBERS",
        "진행", // 등록과 동시에 '진행' 상태로 기록!
        questInfo.score,
        questInfo.method
      ]);
      Logger.log("✅ 오늘 자 자동화 돌발 퀘스트 '" + questInfo.title + "'를 시트에 영구 등기 적재 완료!");
    }
    
    // 3. 전체 회원 쪽지 발송
    var title = preview.title;
    var content = preview.content;
    var res = sendGlobalNotification("quest", title, content);
    
    if (res && res.success) {
      Logger.log("🚀 돌발 퀘스트 자동 전체 알림 등록 및 발송 완료! GLOBAL_ID: " + res.globalId);
    } else {
      Logger.log("⚠️ 돌발 퀘스트 자동 전체 알림 등록 실패: " + (res ? res.error : "알 수 없음"));
    }
  } catch (err) {
    Logger.log("❌ autoSendDailyQuestNotice 오류: " + err.toString());
  }
}

// ──────────────────────────────────────────────
// [이관] autoSendWeeklyRankingNotice
// ──────────────────────────────────────────────
function autoSendWeeklyRankingNotice() {
  try {
    var now = new Date();
    
    // 1. 일회성 트리거 자가 청소 (Clean up one-off trigger)
    var oneOffTriggerId = PropertiesService.getScriptProperties().getProperty("ONE_OFF_WEEKLY_NOTICE_TRIGGER_ID");
    if (oneOffTriggerId) {
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getUniqueId() === oneOffTriggerId) {
          ScriptApp.deleteTrigger(triggers[i]);
          Logger.log("🏆 주간 일회성 트리거 삭제 완료: " + oneOffTriggerId);
          break;
        }
      }
      PropertiesService.getScriptProperties().deleteProperty("ONE_OFF_WEEKLY_NOTICE_TRIGGER_ID");
    }

    // 2. 만약 오늘이 목요일인데 어제(수요일)가 휴일이었다면, 금요일로 연장 예약하고 종료!
    var curDay = now.getDay();
    if (curDay === 4) { // 목요일
      var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
      if (typeof isCenterHoliday === "function" && isCenterHoliday(yesterday)) {
        Logger.log("어제(수요일)가 휴무일이므로 주간 랭킹 발표를 하루 연장합니다. 금요일 새벽 0시 5분에 실행되도록 일회성 트리거를 생성합니다.");
        
        var targetDate = new Date(now.getTime() + 24 * 60 * 60 * 1000);
        targetDate.setHours(0, 5, 0, 0);
        
        var newTrigger = ScriptApp.newTrigger("autoSendWeeklyRankingNotice")
          .timeBased()
          .at(targetDate)
          .create();
        PropertiesService.getScriptProperties().setProperty("ONE_OFF_WEEKLY_NOTICE_TRIGGER_ID", newTrigger.getUniqueId());
        
        // [v67.5] 수요일 휴무로 인한 인바디 마감 연장 안내 쪽지 발송
        var noticeTitle = "📢 [노형빌리지] 주간 인바디 제출 기한 연장 안내 🏋️‍♂️";
        var noticeContent = "안녕하세요, 모험가님!\n\n" +
                            "수요일(어제) 센터 공식 휴무로 인해 이번 주 주간 인바디 측정 및 제출 기한이 오늘(목요일) 밤 23:59까지로 하루 연장되었습니다!\n\n" +
                            "일반 활동 점수(일일 미션 등)는 수요일 자정으로 정상 마감되었으나, 최종 주간 랭킹은 오늘 밤까지 인바디를 제출해 주신 내역까지 포함하여 내일(금요일) 새벽에 정산 및 발표됩니다.\n\n" +
                            "아직 인바디 측정을 완료하지 않으신 모험가님들은 오늘 센터 방문 시 꼭 측정해 주시기 바랍니다!\n\n" +
                            "감사합니다.";
        sendGlobalNotification("quest", noticeTitle, noticeContent);
        
        return; // 목요일 발표 중단
      }
    }

    var preview = getWeeklyRankingNoticePreview();
    if (!preview || !preview.success) {
      Logger.log("주간 명예의 전당 프리뷰 생성 실패로 트리거 종료: " + (preview ? preview.error : ""));
      return;
    }
    
    // 주간 명예의 전당 성적 아카이브 자동 저장
    try {
      var periodMatch = preview.title.match(/\[노형빌리지\] (.+) 주간 명예의 전당/);
      var periodStr = periodMatch ? periodMatch[1] : "";
      if (periodStr) {
        archiveWeeklyRankingToSheet(periodStr);
      }
    } catch (archErr) {
      Logger.log("주간 성적 아카이빙 자동 저장 실패: " + archErr.toString());
    }
    
    var title = preview.title;
    var content = preview.content;
    
    // 단 1회의 전체 알림 적재로 획기적 속도 향상!
    var res = sendGlobalNotification("ranking", title, content);
    if (res && res.success) {
      Logger.log("🏆 주간 명예의 전당 전체 공용 쪽지 단 1개 발송 완료! (행 증가 제로)");
    } else {
      Logger.log("❌ 주간 명예의 전당 전체 쪽지 발송 실패: " + (res ? res.error : "unknown"));
    }
  } catch (err) {
    Logger.log("주간 명예의 전당 자동 쪽지 발송 오류: " + err.toString());
  }
}

// ──────────────────────────────────────────────
// [이관] autoSendMonthlyRankingNotice
// ──────────────────────────────────────────────
function autoSendMonthlyRankingNotice() {
  try {
    var now = new Date();
    
    // 1. 일회성 트리거 자가 청소 (Clean up one-off trigger)
    var oneOffTriggerId = PropertiesService.getScriptProperties().getProperty("ONE_OFF_MONTHLY_NOTICE_TRIGGER_ID");
    if (oneOffTriggerId) {
      var triggers = ScriptApp.getProjectTriggers();
      for (var i = 0; i < triggers.length; i++) {
        if (triggers[i].getUniqueId() === oneOffTriggerId) {
          ScriptApp.deleteTrigger(triggers[i]);
          Logger.log("🏆 월간 일회성 트리거 삭제 완료: " + oneOffTriggerId);
          break;
        }
      }
      PropertiesService.getScriptProperties().deleteProperty("ONE_OFF_MONTHLY_NOTICE_TRIGGER_ID");
    }

    // 2. 만약 오늘이 1일이고 목요일인데 어제(수요일, 이전 달 마지막 날)가 휴일이었다면, 2일로 연장 예약하고 종료!
    var curDate = now.getDate();
    var curDay = now.getDay();
    if (curDate === 1 && curDay === 4) { // 1일이고 목요일인 경우
      var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
      if (typeof isCenterHoliday === "function" && isCenterHoliday(yesterday)) {
        Logger.log("이전 달 마지막 수요일이 휴무일이었고 오늘이 1일(목요일)이므로 월간 랭킹 발표를 하루 연장합니다. 내일(2일) 새벽 0시 10분에 실행되도록 일회성 트리거를 생성합니다.");
        
        var targetDate = new Date(now.getTime() + 24 * 60 * 60 * 1000);
        targetDate.setHours(0, 10, 0, 0);
        
        var newTrigger = ScriptApp.newTrigger("autoSendMonthlyRankingNotice")
          .timeBased()
          .at(targetDate)
          .create();
        PropertiesService.getScriptProperties().setProperty("ONE_OFF_MONTHLY_NOTICE_TRIGGER_ID", newTrigger.getUniqueId());
        
        // [v67.5] 수요일 휴무로 인한 인바디 마감 연장 안내 쪽지 발송
        var noticeTitle = "📢 [노형빌리지] 월간 인바디 제출 기한 연장 안내 🏋️‍♂️";
        var noticeContent = "안녕하세요, 모험가님!\n\n" +
                            "월말 수요일(어제) 센터 공식 휴무로 인해 이번 달 월간 최종 인바디 측정 및 제출 기한이 오늘(목요일) 밤 23:59까지로 하루 연장되었습니다!\n\n" +
                            "일반 활동 점수(일일 미션 등)는 월말일 자정으로 정상 마감되었으나, 최종 월간 MVP 선정은 오늘 밤까지 인바디를 제출해 주신 내역까지 포함하여 내일(금요일) 새벽에 정산 및 발표됩니다.\n\n" +
                            "아직 인바디 측정을 완료하지 않으신 모험가님들은 오늘 센터 방문 시 꼭 측정해 주시기 바랍니다!\n\n" +
                            "감사합니다.";
        sendGlobalNotification("quest", noticeTitle, noticeContent);
        
        return; // 1일 발표 중단
      }
    }

    var preview = getMonthlyRankingNoticePreview();
    if (!preview || !preview.success) {
      Logger.log("월간 명예의 전당 프리뷰 생성 실패로 트리거 종료: " + (preview ? preview.error : ""));
      return;
    }
    
    // 월간 명예의 전당 성적 아카이브 자동 저장
    try {
      var periodMatch = preview.title.match(/\[노형빌리지\] (.+) 월간 명예의 전당/);
      var periodStr = periodMatch ? periodMatch[1] : "";
      if (periodStr) {
        archiveMonthlyRankingToSheet(periodStr);
      }
    } catch (archErr) {
      Logger.log("월간 성적 아카이빙 자동 저장 실패: " + archErr.toString());
    }
    
    var title = preview.title;
    var content = preview.content;
    
    // 단 1회의 전체 알림 적재로 획기적 속도 향상!
    var res = sendGlobalNotification("ranking", title, content);
    if (res && res.success) {
      Logger.log("🏆 월간 명예의 전당 전체 공용 쪽지 단 1개 발송 완료! (행 증가 제로)");
    } else {
      Logger.log("❌ 월간 명예의 전당 전체 쪽지 발송 실패: " + (res ? res.error : "unknown"));
    }
  } catch (err) {
    Logger.log("월간 명예의 전당 자동 쪽지 발송 오류: " + err.toString());
  }
}

// ──────────────────────────────────────────────
// [이관] getWeeklyRankingNoticePreview
// ──────────────────────────────────────────────
function getWeeklyRankingNoticePreview() {
  try {
    var hofRes = getHallOfFameData();
    if (!hofRes || !hofRes.success || !hofRes.data || !hofRes.data.archive || !hofRes.data.archive.weekly) {
      return { success: false, error: "주간 명예의 전당 데이터 취합 실패" };
    }
    
    var weeklyArchive = hofRes.data.archive.weekly;
    if (weeklyArchive.length === 0) {
      return { success: false, error: "과거 주간 랭킹 아카이브 데이터가 없습니다." };
    }
    
    // 가장 최근 완성된 주차 (직전 주)
    var latestWeek = weeklyArchive[0];
    var period = latestWeek.period; // 예: "5월 4주"
    
    // winners parsing ("🥇 1.김순희(344p) | 2.박노형(320p)...")
    var winnersStr = latestWeek.winners.replace("🥇 ", "");
    var tokens = winnersStr.split("|");
    var w1 = tokens[0] ? tokens[0].trim() : "-";
    var w2 = tokens[1] ? tokens[1].trim() : "-";
    var w3 = tokens[2] ? tokens[2].trim() : "-";
    
    var title = "🏆 [노형빌리지] " + period + " 주간 명예의 전당 발표! ✉️";
    var content = "🏰 [노형빌리지 " + period + " 주간 시상식] 🏰\n\n" +
                  "지난 한 주 동안 노형빌리지를 뜨겁게 빛내주신 위대한 영웅들의 주간 명예의 전당 순위가 발표되었습니다!\n\n" +
                  "👑 [최종 주간 TOP 3]\n" +
                  "🥇 1위 : " + w1 + "\n" +
                  "🥈 2위 : " + w2 + "\n" +
                  "🥉 3위 : " + w3 + "\n\n" +
                  "수상하신 모든 모험가님 진심으로 축하드립니다! 👏\n\n" +
                  "우편함의 바로가기 버튼을 터치하여 전체 10위까지의 영광스러운 전체 순위표를 확인하고 축하의 박수를 보내주세요! 🎉";
                  
    return { success: true, title: title, content: content };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ──────────────────────────────────────────────
// [이관] getMonthlyRankingNoticePreview
// ──────────────────────────────────────────────
function getMonthlyRankingNoticePreview() {
  try {
    var hofRes = getHallOfFameData();
    if (!hofRes || !hofRes.success || !hofRes.data || !hofRes.data.archive) {
      return { success: false, error: "월간 명예의 전당 데이터 취합 실패" };
    }
    
    var monthlyArchive = hofRes.data.archive.monthly || [];
    var attArchive = hofRes.data.archive.attendance || [];
    
    if (monthlyArchive.length === 0 && attArchive.length === 0) {
      return { success: false, error: "과거 월간/출석왕 아카이브 데이터가 없습니다." };
    }
    
    // 가장 최근 완성된 월 (직전 월)
    var latestMonth = monthlyArchive[0];
    var latestAtt = attArchive[0];
    
    var displayPeriod = "";
    if (latestMonth) displayPeriod = latestMonth.period;
    else if (latestAtt) displayPeriod = latestAtt.period;
    
    var w1 = "-", w2 = "-", w3 = "-";
    if (latestMonth) {
      var mWinnersStr = latestMonth.winners.replace("🥇 ", "");
      var mTokens = mWinnersStr.split("|");
      w1 = mTokens[0] ? mTokens[0].trim() : "-";
      w2 = mTokens[1] ? mTokens[1].trim() : "-";
      w3 = mTokens[2] ? mTokens[2].trim() : "-";
    }
    
    var att1 = "-", att2 = "-", att3 = "-";
    if (latestAtt) {
      var attTokens = latestAtt.winners.split("|");
      att1 = attTokens[0] ? attTokens[0].trim() : "-";
      att2 = attTokens[1] ? attTokens[1].trim() : "-";
      att3 = attTokens[2] ? attTokens[2].trim() : "-";
    }
    
    var title = "📅 [노형빌리지] " + displayPeriod + " 월간 명예의 전당 발표! ✉️";
    var content = "🏰 [노형빌리지 " + displayPeriod + " 월간 & 출석왕 시상식] 🏰\n\n" +
                  "지난 한 달 동안 노형빌리지를 가장 뜨겁게 불태워주신 명예로운 월간 MVP와 출석왕이 탄생했습니다!\n\n" +
                  "🔥 [최종 월간 TOP 3]\n" +
                  "🥇 1위 : " + w1 + "\n" +
                  "🥈 2위 : " + w2 + "\n" +
                  "🥉 3위 : " + w3 + "\n\n" +
                  "🏃 [최종 월별 출석왕 TOP 3]\n" +
                  att1 + "\n" +
                  att2 + "\n" +
                  att3 + "\n\n" +
                  "위대한 성과를 거두신 모든 영웅분들께 뜨거운 박수를 보냅니다! 👏\n\n" +
                  "우편함의 바로가기 버튼을 터치하여 전체 명예의 전당에서 영광스러운 순위표를 즉시 확인해 보세요! 💫";
                  
    return { success: true, title: title, content: content };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ──────────────────────────────────────────────
// [이관] archiveWeeklyRankingToSheet
// ──────────────────────────────────────────────
function archiveWeeklyRankingToSheet(period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지_주간성적아카이브") || ss.insertSheet("33챌린지_주간성적아카이브");
    
    if (sheet.getLastRow() === 0) {
      var headers = ["주차", "순위", "이름", "연락처", "체력", "실천력", "회복력", "주간토탈점수", "누적토탈점수"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f1f3f5");
    }
    
    // 중복 제거 검사
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === period.trim()) {
        Logger.log(period + " 주간 기록은 이미 아카이브에 존재하므로 건너뜁니다.");
        return { success: true, message: "이미 기록이 존재합니다." };
      }
    }
    
    // calculateWeeklyRankingForPeriod 호출하여 정확한 감량/체력 점수 및 총점이 합산된 정렬 순위를 가져옴!
    var result = calculateWeeklyRankingForPeriod(period);
    if (!result || !result.rankings || result.rankings.length === 0) {
      return { success: false, message: "해당 주차의 데이터를 계산하지 못했습니다." };
    }
    
    var rankedUsers = result.rankings;
    var rowsToAppend = [];
    for (var idx = 0; idx < rankedUsers.length; idx++) {
      var u = rankedUsers[idx];
      rowsToAppend.push([
        period,
        idx + 1,
        u.name,
        "'" + u.phone,
        u.health,
        u.perf,
        u.def,
        u.score, // weeklyTotal (주간토탈점수)
        u.lifetimeTotal
      ]);
    }
    
    if (rowsToAppend.length > 0) {
      sheet.insertRowsAfter(1, rowsToAppend.length);
      sheet.getRange(2, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      Logger.log(period + " 주간 성적 아카이브 최신 탑 적재 완료: 총 " + rowsToAppend.length + "행");
    }
    return { success: true, message: "아카이브 저장 성공" };
  } catch (e) {
    Logger.log("archiveWeeklyRankingToSheet 오류: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ──────────────────────────────────────────────
// [이관] archiveMonthlyRankingToSheet
// ──────────────────────────────────────────────
function archiveMonthlyRankingToSheet(period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지_월간성적아카이브") || ss.insertSheet("33챌린지_월간성적아카이브");
    
    if (sheet.getLastRow() === 0 || sheet.getLastColumn() < 10) {
      var headers = ["월", "순위", "이름", "연락처", "체력", "실천력", "회복력", "월간토탈점수", "누적토탈점수", "출석기록"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f1f3f5");
    }
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === period.trim()) {
        Logger.log(period + " 월간 기록은 이미 아카이브에 존재하므로 스킵합니다.");
        return { success: true, message: "이미 기록이 존재합니다." };
      }
    }
    
    // calculateMonthlyRankingForPeriod 호출하여 정확한 감량/체력 점수 및 총점이 합산된 정렬 순위를 가져옴!
    var result = calculateMonthlyRankingForPeriod(period);
    if (!result || !result.rankings || result.rankings.length === 0) {
      return { success: false, message: "해당 월의 데이터를 계산하지 못했습니다." };
    }
    
    var rankedUsers = result.rankings;
    var rowsToAppend = [];
    for (var idx = 0; idx < rankedUsers.length; idx++) {
      var u = rankedUsers[idx];
      rowsToAppend.push([
        period,
        idx + 1,
        u.name,
        "'" + u.phone,
        u.health,
        u.perf,
        u.def,
        u.score, // monthlyTotal (월간토탈점수)
        u.lifetimeTotal,
        u.attCount
      ]);
    }
    
    if (rowsToAppend.length > 0) {
      sheet.insertRowsAfter(1, rowsToAppend.length);
      sheet.getRange(2, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
      Logger.log(period + " 월간 성적 아카이브 최신 탑 적재 완료: 총 " + rowsToAppend.length + "행");
    }
    return { success: true, message: "아카이브 저장 성공" };
  } catch (e) {
    Logger.log("archiveMonthlyRankingToSheet 오류: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ──────────────────────────────────────────────
// [이관] calculateWeeklyRankingForPeriod
// ──────────────────────────────────────────────
function calculateWeeklyRankingForPeriod(period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = ss.getSheetByName("일일_활동_기록");
    var summaryData = summarySheet ? summarySheet.getDataRange().getValues() : [];
    
    var targetThuDate = null;
    for (var j = 1; j < summaryData.length; j++) {
      var recDateRaw = summaryData[j][0];
      var recDate = (recDateRaw instanceof Date) ? recDateRaw : new Date(recDateRaw);
      if (isNaN(recDate.getTime())) continue;
      
      if (getWeekStringLocal(recDate) === period) {
        targetThuDate = getThuStartOfWeekLocal(recDate);
        break;
      }
    }
    
    if (!targetThuDate) {
      targetThuDate = getThuStartOfWeekLocal(new Date());
    }
    
    var endOfActivityDate = new Date(targetThuDate.getTime() + 6 * 24 * 60 * 60 * 1000 + 23 * 60 * 60 * 1000 + 59 * 60 * 1000 + 59 * 1000); // 수요일 23:59:59
    var endOfInbodyDate = new Date(endOfActivityDate.getTime()); // 기본은 수요일 23:59:59
    
    // [v67.5] 수요일이 공휴일/휴무일이면 목요일 밤 23:59:59로 인바디 마감만 연장!
    var lastWed = new Date(targetThuDate.getTime() + 6 * 24 * 60 * 60 * 1000);
    if (typeof isCenterHoliday === "function" && isCenterHoliday(lastWed)) {
      endOfInbodyDate.setDate(endOfInbodyDate.getDate() + 1);
    }
    
    // 인바디 로드
    var inbodySheet = ss.getSheetByName("33챌린지_인바디");
    var inData = inbodySheet ? inbodySheet.getDataRange().getValues() : [];
    var userInbodyMap = {};
    for (var k = 1; k < inData.length; k++) {
      var rawPhone = String(inData[k][2] || "").trim();
      var phone = formatPhoneNumber(rawPhone).replace(/[^0-9]/g, "");
      if (!phone) continue;
      userInbodyMap[phone] = userInbodyMap[phone] || [];
      userInbodyMap[phone].push({
        date: new Date(inData[k][0]),
        weight: Number(inData[k][3] || 0),
        muscle: Number(inData[k][4] || 0),
        fat: Number(inData[k][5] || 0)
      });
    }
    
    // 회원명단에서 목표체중 로드
    var userTargetWeightMap = {};
    var memberSheet = ss.getSheetByName("회원명단");
    if (memberSheet) {
      var mData = memberSheet.getDataRange().getDisplayValues();
      var mCols = getMemberSheetColumnIndices(memberSheet);
      for (var i = 1; i < mData.length; i++) {
        var mPhone = normalizePhoneDigits(mData[i][mCols.phone]);
        if (mPhone) {
          userTargetWeightMap[mPhone] = parseWeightSafely(mData[i][mCols.targetWeight]);
        }
      }
    }
    
    // 일일 점수 누적 집계
    var userStats = {};
    for (var j = 1; j < summaryData.length; j++) {
      var rawPhone = String(summaryData[j][1] || "").trim();
      var phone = formatPhoneNumber(rawPhone).replace(/[^0-9]/g, "");
      if (!phone) continue;
      
      var recDateRaw = summaryData[j][0];
      var recDate = (recDateRaw instanceof Date) ? recDateRaw : new Date(recDateRaw);
      if (isNaN(recDate.getTime())) continue;
      
      if (recDate <= endOfActivityDate) {
        var name = String(summaryData[j][2] || "모험가").trim();
        var visitScore = Number(summaryData[j][3] || 0);
        var intensityScore = Number(summaryData[j][4] || 0);
        var actScore = Number(summaryData[j][5] || 0);
        var recScore = Number(summaryData[j][6] || 0);
        var hpScore = Number(summaryData[j][7] || 0);
        var rowTotal = Number(summaryData[j][9] || 0);
        
        if (!userStats[phone]) {
          userStats[phone] = { name: name, health: 0, perf: 0, def: 0, weeklyTotalAct: 0, lifetimeTotalAct: 0 };
        }
        
        var stat = userStats[phone];
        stat.lifetimeTotalAct += rowTotal;
        
        if (recDate >= targetThuDate && recDate <= endOfActivityDate) {
          stat.health += hpScore;
          stat.perf += (visitScore + intensityScore + actScore);
          stat.def += recScore;
          stat.weeklyTotalAct += rowTotal;
        }
      }
    }
    
    var list = [];
    for (var phone in userStats) {
      var stat = userStats[phone];
      var records = userInbodyMap[phone] || [];
      
      var firstEver = null;
      var latestEver = null;
      var latestInWeek = null;
      for (var rIdx = 0; rIdx < records.length; rIdx++) {
        var r = records[rIdx];
        if (r.date <= endOfInbodyDate) {
          if (!firstEver || r.date < firstEver.date) firstEver = r;
          if (!latestEver || r.date > latestEver.date) latestEver = r;
          if (r.date >= targetThuDate && r.date <= endOfInbodyDate) {
            if (!latestInWeek || r.date > latestInWeek.date) latestInWeek = r;
          }
        }
      }
      
      var targetWeight = userTargetWeightMap[phone] || 0;
      var inbodyWeeklyScore = 0;
      if (latestInWeek) {
        var prevBeforeThisWeek = null;
        for (var rIdx = 0; rIdx < records.length; rIdx++) {
          var r = records[rIdx];
          if (r.date < targetThuDate) {
            if (!prevBeforeThisWeek || r.date > prevBeforeThisWeek.date) prevBeforeThisWeek = r;
          }
        }
        var baseRecord = prevBeforeThisWeek || firstEver;
        inbodyWeeklyScore = calculateInbodyScoreHelper(baseRecord, latestInWeek, targetWeight, 'weekly');
      }
      
      var inbodyLifetimeScore = calculateInbodyScoreHelper(firstEver, latestEver, targetWeight, 'lifetime');
      
      var totalWeeklyExp = stat.weeklyTotalAct + inbodyWeeklyScore;
      var totalLifetimeExp = stat.lifetimeTotalAct + inbodyLifetimeScore;
      
      // 인바디 감량 점수 및 보너스 점수 등은 모두 체력으로 들어갑니다!
      var healthScore = stat.health + inbodyWeeklyScore;
      
      list.push({
        name: stat.name,
        phone: phone,
        health: healthScore,
        perf: stat.perf,
        def: stat.def,
        score: totalWeeklyExp,
        lifetimeTotal: totalLifetimeExp
      });
    }
    
    list.sort(function(a, b) { return b.score - a.score; });
    for (var idx = 0; idx < list.length; idx++) {
      list[idx].rank = idx + 1;
    }
    return { period: period, rankings: list };
  } catch (e) {
    Logger.log("calculateWeeklyRankingForPeriod 에러: " + e.toString());
    return null;
  }
}

// ──────────────────────────────────────────────
// [이관] calculateMonthlyRankingForPeriod
// ──────────────────────────────────────────────
function calculateMonthlyRankingForPeriod(period) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = ss.getSheetByName("일일_활동_기록");
    var summaryData = summarySheet ? summarySheet.getDataRange().getValues() : [];
    
    var targetYear = new Date().getFullYear();
    var monthNum = parseInt(period.replace("월", "").trim());
    if (isNaN(monthNum)) monthNum = new Date().getMonth() + 1;
    
    var startOfMonth = new Date(targetYear, monthNum - 1, 1, 0, 0, 0, 0);
    var endOfMonth = new Date(targetYear, monthNum, 0, 23, 59, 59, 999);
    
    // 월초 인바디 마감 기한
    var firstDay = new Date(targetYear, monthNum - 1, 1);
    var firstDayOfWeek = firstDay.getDay();
    var diffToWed = (3 - firstDayOfWeek + 7) % 7;
    var baselineDeadline = new Date(targetYear, monthNum - 1, 1 + diffToWed);
    var hasWedHoliday = isCenterHoliday(baselineDeadline);
    if (hasWedHoliday) {
      baselineDeadline.setDate(baselineDeadline.getDate() + 1);
    }
    baselineDeadline.setHours(23, 59, 59, 999);
    
    // 인바디 로드
    var inbodySheet = ss.getSheetByName("33챌린지_인바디");
    var inData = inbodySheet ? inbodySheet.getDataRange().getValues() : [];
    var userInbodyMap = {};
    for (var k = 1; k < inData.length; k++) {
      var rawPhone = String(inData[k][2] || "").trim();
      var phone = formatPhoneNumber(rawPhone).replace(/[^0-9]/g, "");
      if (!phone) continue;
      userInbodyMap[phone] = userInbodyMap[phone] || [];
      userInbodyMap[phone].push({
        date: new Date(inData[k][0]),
        weight: Number(inData[k][3] || 0),
        muscle: Number(inData[k][4] || 0),
        fat: Number(inData[k][5] || 0)
      });
    }
    
    // 회원명단에서 목표체중 로드
    var userTargetWeightMap = {};
    var memberSheet = ss.getSheetByName("회원명단");
    if (memberSheet) {
      var mData = memberSheet.getDataRange().getDisplayValues();
      var mCols = getMemberSheetColumnIndices(memberSheet);
      for (var i = 1; i < mData.length; i++) {
        var mPhone = normalizePhoneDigits(mData[i][mCols.phone]);
        if (mPhone) {
          userTargetWeightMap[mPhone] = parseWeightSafely(mData[i][mCols.targetWeight]);
        }
      }
    }
    
    // 일일 점수 누적 집계
    var userStats = {};
    for (var j = 1; j < summaryData.length; j++) {
      var rawPhone = String(summaryData[j][1] || "").trim();
      var phone = formatPhoneNumber(rawPhone).replace(/[^0-9]/g, "");
      if (!phone) continue;
      
      var recDateRaw = summaryData[j][0];
      var recDate = (recDateRaw instanceof Date) ? recDateRaw : new Date(recDateRaw);
      if (isNaN(recDate.getTime())) continue;
      
      if (recDate <= endOfMonth) {
        var name = String(summaryData[j][2] || "모험가").trim();
        var visitScore = Number(summaryData[j][3] || 0);
        var intensityScore = Number(summaryData[j][4] || 0);
        var actScore = Number(summaryData[j][5] || 0);
        var recScore = Number(summaryData[j][6] || 0);
        var hpScore = Number(summaryData[j][7] || 0);
        var rowTotal = Number(summaryData[j][9] || 0);
        
        if (!userStats[phone]) {
          userStats[phone] = { name: name, health: 0, perf: 0, def: 0, monthlyTotalAct: 0, lifetimeTotalAct: 0 };
        }
        
        var stat = userStats[phone];
        stat.lifetimeTotalAct += rowTotal;
        
        if (recDate >= startOfMonth && recDate <= endOfMonth) {
          stat.health += hpScore;
          stat.perf += (visitScore + intensityScore + actScore);
          stat.def += recScore;
          stat.monthlyTotalAct += rowTotal;
        }
      }
    }
    
    var list = [];
    for (var phone in userStats) {
      var stat = userStats[phone];
      var records = userInbodyMap[phone] || [];
      
      var firstEver = null;
      var latestEver = null;
      var baselineRecord = null;
      var latestInMonth = null;
      for (var rIdx = 0; rIdx < records.length; rIdx++) {
        var r = records[rIdx];
        if (r.date <= endOfMonth) {
          if (!firstEver || r.date < firstEver.date) firstEver = r;
          if (!latestEver || r.date > latestEver.date) latestEver = r;
          
          if (r.date >= startOfMonth && r.date <= baselineDeadline) {
            var rDay = r.date.getDay();
            var isValidDay = (rDay >= 1 && rDay <= 3) || (hasWedHoliday && rDay === 4);
            if (isValidDay) {
              if (!baselineRecord || r.date < baselineRecord.date) baselineRecord = r;
            }
          }
          if (r.date >= startOfMonth && isDateInLastWeekMonToWed(r.date)) {
            if (!latestInMonth || r.date > latestInMonth.date) latestInMonth = r;
          }
        }
      }
      
      var targetWeight = userTargetWeightMap[phone] || 0;
      var inbodyMonthlyScore = 0;
      if (baselineRecord && latestInMonth) {
        inbodyMonthlyScore = calculateInbodyScoreHelper(baselineRecord, latestInMonth, targetWeight, 'monthly');
      }
      
      var inbodyLifetimeScore = calculateInbodyScoreHelper(firstEver, latestEver, targetWeight, 'lifetime');
      
      var totalMonthlyExp = stat.monthlyTotalAct + inbodyMonthlyScore;
      var totalLifetimeExp = stat.lifetimeTotalAct + inbodyLifetimeScore;
      
      // 인바디 감량 점수 및 명품 유지 보너스 등은 모두 체력으로 합산됩니다!
      var healthScore = stat.health + inbodyMonthlyScore;
      
      list.push({
        name: stat.name,
        phone: phone,
        health: healthScore,
        perf: stat.perf,
        def: stat.def,
        score: totalMonthlyExp,
        lifetimeTotal: totalLifetimeExp
      });
    }
    
    list.sort(function(a, b) { return b.score - a.score; });
    for (var idx = 0; idx < list.length; idx++) {
      list[idx].rank = idx + 1;
    }
    
    // 출석 횟수 집계
    var monthlyAttendanceMap = {};
    try {
      var attSheet = ss.getSheetByName("출석기록");
      var attData = attSheet ? attSheet.getDataRange().getValues() : [];
      if (attSheet && attData.length > 1) {
        var attCols = getAttendanceColumnIndices(attSheet);
        for (var aIdx = 1; aIdx < attData.length; aIdx++) {
          var aRow = attData[aIdx];
          var aPhoneRaw = String(aRow[attCols.phone] || "").trim();
          var aPhone = formatPhoneNumber(aPhoneRaw).replace(/[^0-9]/g, "");
          if (!aPhone) continue;
          
          var aDateRaw = aRow[attCols.date];
          var aDate = (aDateRaw instanceof Date) ? aDateRaw : new Date(aDateRaw);
          if (isNaN(aDate.getTime())) continue;
          
          if (aDate >= startOfMonth && aDate <= endOfMonth) {
            monthlyAttendanceMap[aPhone] = (monthlyAttendanceMap[aPhone] || 0) + 1;
          }
        }
      }
    } catch (attErr) {
      Logger.log("출석 횟수 실시간 집계 중 오류: " + attErr.toString());
    }
    
    list.forEach(function(u) {
      u.attCount = monthlyAttendanceMap[u.phone] || 0;
    });
    
    return { period: period, rankings: list };
  } catch (e) {
    Logger.log("calculateMonthlyRankingForPeriod 에러: " + e.toString());
    return null;
  }
}

// ──────────────────────────────────────────────
// [이관] migrateSheetPeriods
// ──────────────────────────────────────────────
function migrateSheetPeriods() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지_주간성적아카이브");
    if (!sheet) {
      Logger.log("33챌린지_주간성적아카이브 시트를 찾을 수 없습니다.");
      return;
    }
    
    // 1. 기존 과거 주차 이름 보정 및 지난주 버그 행 일괄 삭제 (역순 처리로 밀림 방지)
    var range = sheet.getDataRange();
    var values = range.getValues();
    var count = 0;
    
    for (var i = values.length - 1; i >= 1; i--) {
      var period = String(values[i][0]).trim();
      var rowNum = i + 1;
      
      if (period === "5월 2주") {
        sheet.getRange(rowNum, 1).setValue("5월 3주");
        count++;
      } else if (period === "5월 1주") {
        sheet.getRange(rowNum, 1).setValue("5월 2주");
        count++;
      } else if (period === "4월 5주") {
        sheet.getRange(rowNum, 1).setValue("5월 1주");
        count++;
      } else if (period === "5월 3주" || period === "5월 4주") {
        // [체력 0점 버그 해결] 지난주(5월 21일~27일) 기록은 옛날 버그 코드로 잘못 적힌 행들이므로 삭제하고 새로 씁니다!
        sheet.deleteRow(rowNum);
      }
    }
    
    // 2. 지난주(5월 21일~27일, 즉 "5월 4주") 성적을 정확한 점수 계산법(인바디 및 보너스가 합산된 체력점수 포함!)으로 실시간 재계산해서 깔끔히 재적재!
    var targetPeriod = "5월 4주";
    var result = calculateWeeklyRankingForPeriod(targetPeriod);
    if (result && result.rankings && result.rankings.length > 0) {
      var rankedUsers = result.rankings;
      var rowsToAppend = [];
      for (var idx = 0; idx < rankedUsers.length; idx++) {
        var u = rankedUsers[idx];
        rowsToAppend.push([
          targetPeriod,
          idx + 1,
          u.name,
          "'" + u.phone,
          u.health,
          u.perf,
          u.def,
          u.score, // weeklyTotal (주간토탈)
          u.lifetimeTotal
        ]);
      }
      
      if (rowsToAppend.length > 0) {
        sheet.insertRowsAfter(1, rowsToAppend.length);
        sheet.getRange(2, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
        count += rowsToAppend.length;
        Logger.log("지난주 5월 4주차 성적(체력점수 정상 합산본)을 신규 마이그레이션 적재하였습니다.");
      }
    }
    
    Logger.log("마이그레이션 최종 완료: 총 " + count + "행의 데이터 및 주차명을 성공적으로 올바르게 보정했습니다.");
  } catch (e) {
    Logger.log("마이그레이션 에러: " + e.toString());
  }
}

// ──────────────────────────────────────────────
// [이관] rollbackPrematureJune1stWeekRanking
// ──────────────────────────────────────────────
function rollbackPrematureJune1stWeekRanking() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 33챌린지_주간성적아카이브 시트에서 "6월 1주" 행 삭제
  var archiveSheet = ss.getSheetByName("33챌린지_주간성적아카이브");
  if (archiveSheet) {
    var data = archiveSheet.getDataRange().getValues();
    var deletedCount = 0;
    // 아래서부터 위로 삭제해야 인덱스가 꼬이지 않음
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === "6월 1주") {
        archiveSheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    Logger.log("주간성적아카이브에서 '6월 1주' 행 " + deletedCount + "개 삭제 완료.");
  }
  
  // 2. 전체_알림_목록 시트에서 "6월 1주 주간 명예의 전당 발표" 행 삭제
  var alarmSheet = ss.getSheetByName("전체_알림_목록") || ss.getSheetByName("전체알림");
  if (alarmSheet) {
    var data = alarmSheet.getDataRange().getValues();
    var deletedAlarms = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      var title = String(data[i][3] || "");
      if (title.indexOf("6월 1주 주간 명예의 전당 발표") !== -1) {
        alarmSheet.deleteRow(i + 1);
        deletedAlarms++;
      }
    }
    Logger.log("전체_알림_목록에서 '6월 1주' 발표 쪽지 " + deletedAlarms + "개 삭제 완료.");
  }
  
  // 3. 내일(금요일, 6월 5일) 새벽 0시 5분에 발표가 자동 실행되도록 일회성 트리거 수동 생성
  var oneOffTriggerId = PropertiesService.getScriptProperties().getProperty("ONE_OFF_WEEKLY_NOTICE_TRIGGER_ID");
  if (oneOffTriggerId) {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      if (triggers[i].getUniqueId() === oneOffTriggerId) {
        ScriptApp.deleteTrigger(triggers[i]);
        break;
      }
    }
  }
  
  var targetDate = new Date(2026, 5, 5, 0, 5, 0); // 2026년 6월 5일 00:05 (KST)
  var newTrigger = ScriptApp.newTrigger("autoSendWeeklyRankingNotice")
    .timeBased()
    .at(targetDate)
    .create();
  PropertiesService.getScriptProperties().setProperty("ONE_OFF_WEEKLY_NOTICE_TRIGGER_ID", newTrigger.getUniqueId());
  
  // 캐시 초기화
  var cache = CacheService.getScriptCache();
  cache.remove("v45_member_registry");
  
  Logger.log("6월 5일(금) 새벽 0시 5분에 주간 랭킹을 정상 재정산 및 재발표하도록 일회성 트리거를 정상 예약 완료했습니다!");
}

// ──────────────────────────────────────────────
// [이관] runBackendMigration
// ──────────────────────────────────────────────
function runBackendMigration() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("지혜의_보물고");
    if (!sheet) return { success: false, error: "지혜의_보물고 시트가 존재하지 않습니다." };
    
    var pKeyRes = getPollinationsApiKey();
    var pKey = pKeyRes.success ? pKeyRes.key : "";
    if (!pKey) {
      return { success: false, error: "환경설정 시트 B7 셀에 Pollinations API Key(sk_...)를 입력해 주세요." };
    }
    
    var data = sheet.getDataRange().getValues();
    var successCount = 0;
    var failCount = 0;
    
    var folder = getOrCreateWisdomFolder();
    
    for (var i = 1; i < data.length; i++) {
      var rawImage = String(data[i][8] || "").trim(); // I열
      if (rawImage && rawImage.indexOf("pollinations.ai") > -1) {
        try {
          var cleanUrl = rawImage.replace(/[?&]key=[^&]*/g, "");
          if (cleanUrl.indexOf("?") === -1) {
            cleanUrl += "?";
          } else if (!cleanUrl.endsWith("&") && !cleanUrl.endsWith("?")) {
            cleanUrl += "&";
          }
          
          if (cleanUrl.indexOf("nologo=") === -1) {
            cleanUrl += "nologo=true&private=true&feed=false&";
          }
          cleanUrl += "key=" + pKey;
          
          Logger.log("🔄 Fetching image: " + cleanUrl);
          
          var response = UrlFetchApp.fetch(cleanUrl, {
            muteHttpExceptions: true
          });
          
          if (response.getResponseCode() !== 200) {
            throw new Error("HTTP " + response.getResponseCode() + ": " + response.getContentText());
          }
          
          var blob = response.getBlob();
          var contentType = blob.getContentType();
          if (contentType.indexOf("image") === -1) {
            blob.setContentType("image/png");
          }
          
          var timestamp = Date.now();
          blob.setName("wisdom_migrated_" + (i + 1) + "_" + timestamp + ".png");
          
          var file = folder.createFile(blob);
          var driveUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
          
          sheet.getRange(i + 1, 9).setValue(driveUrl);
          successCount++;
          Logger.log("✅ Row " + (i + 1) + " migrated to: " + driveUrl);
        } catch (e) {
          failCount++;
          Logger.log("❌ Row " + (i + 1) + " migration failed: " + e.toString());
        }
        
        Utilities.sleep(500); // 0.5초 대기
      }
    }
    
    return { success: true, successCount: successCount, failCount: failCount };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ──────────────────────────────────────────────
// [이관] backfillAllRankingArchives
// ──────────────────────────────────────────────
function backfillAllRankingArchives() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = ss.getSheetByName("일일_활동_기록");
    if (!summarySheet) return { success: false, error: "일일_활동_기록 시트가 없습니다." };
    
    var summaryData = summarySheet.getDataRange().getValues();
    var now = new Date();
    
    

    var startOfWeek = getThuStartOfWeekLocal(now);
    var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
    
    // (A) 과거 고유 주차 리스트 추출
    var uniqueWeeks = {};
    for (var j = 1; j < summaryData.length; j++) {
      var recDateRaw = summaryData[j][0];
      var recDate = (recDateRaw instanceof Date) ? recDateRaw : new Date(recDateRaw);
      if (isNaN(recDate.getTime())) continue;
      
      if (recDate < startOfWeek) {
        var weekKey = getWeekStringLocal(recDate);
        uniqueWeeks[weekKey] = getThuStartOfWeekLocal(recDate);
      }
    }
    var weekKeys = Object.keys(uniqueWeeks);
    weekKeys.sort(function(a, b) { return uniqueWeeks[a] - uniqueWeeks[b]; });
    
    var weeklyCount = 0;
    for (var i = 0; i < weekKeys.length; i++) {
      var res = archiveWeeklyRankingToSheet(weekKeys[i]);
      if (res && res.success) weeklyCount++;
    }
    
    // (B) 과거 고유 월 리스트 추출
    var uniqueMonths = {};
    for (var j = 1; j < summaryData.length; j++) {
      var recDateRaw = summaryData[j][0];
      var recDate = (recDateRaw instanceof Date) ? recDateRaw : new Date(recDateRaw);
      if (isNaN(recDate.getTime())) continue;
      
      if (recDate < startOfMonth) {
        var monthKey = (recDate.getMonth() + 1) + "월";
        var sortDate = new Date(recDate.getFullYear(), recDate.getMonth(), 1);
        uniqueMonths[monthKey] = sortDate;
      }
    }
    var monthKeys = Object.keys(uniqueMonths);
    monthKeys.sort(function(a, b) { return uniqueMonths[a] - uniqueMonths[b]; });
    
    var monthlyCount = 0;
    for (var i = 0; i < monthKeys.length; i++) {
      var res = archiveMonthlyRankingToSheet(monthKeys[i]);
      if (res && res.success) monthlyCount++;
    }
    
    return { 
      success: true, 
      message: "백필 완료! 주간: " + weekKeys.length + "개 중 " + weeklyCount + "개 완료 | 월간: " + monthKeys.length + "개 중 " + monthlyCount + "개 완료." 
    };
  } catch (e) {
    Logger.log("backfillAllRankingArchives 오류: " + e.toString());
    return { success: false, error: e.toString() };
  }
}