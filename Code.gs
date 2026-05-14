/**
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * v1.0 - Core Routing & Database Setup
 */

function getArchiveFeed() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("아카이브");
    if (!sheet) return { error: "'아카이브' 시트를 찾을 수 없습니다." };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return []; 
    
    var startRow = Math.max(2, lastRow - 19);
    var numRows = lastRow - startRow + 1;
    var data = sheet.getRange(startRow, 1, numRows, 9).getValues();
    
    return data.reverse().map(function(row) {
      return {
        date: row[0] instanceof Date ? Utilities.formatDate(row[0], "GMT+9", "yyyy-MM-dd") : String(row[0]),
        time: row[1] instanceof Date ? Utilities.formatDate(row[1], "GMT+9", "HH:mm:ss") : String(row[1]),
        name: String(row[2] || ""),
        type: String(row[4] || ""),
        item: String(row[5] || ""),
        comment: String(row[6] || ""),
        photoId: String(row[7] || ""),
        score: row[8]
      };
    });
  } catch (e) {
    return { error: "피드 로딩 실패: " + e.toString() };
  }
}

/**
 * [아카이브] 진행 중인 돌발 퀘스트 가져오기
 */
function getActiveEvents() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("설정");
    if (!configSheet) return [];
    var data = configSheet.getDataRange().getValues();
    var events = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "돌발퀘스트" && data[i][2] === "진행중") {
        events.push({ title: data[i][1] });
      }
    }
    return events;
  } catch (e) { return []; }
}

/**
 * [아카이브] 인증 기록 제출
 */
function submitArchive(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("아카이브") || ss.insertSheet("아카이브");
    
    if (sheet.getLastRow() === 0) {
      var arHeaders = ["날짜", "시간", "이름", "전화번호", "유형", "항목", "코멘트", "사진ID", "점수"];
      sheet.getRange(1, 1, 1, arHeaders.length).setValues([arHeaders]);
    }

    var now = new Date();
    var photoId = "";
    
    if (payload.image && payload.image.indexOf(",") > -1) {
      try {
        var folderName = "GenieWorld_Archive";
        var folders = DriveApp.getFoldersByName(folderName);
        var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        
        var base64Data = payload.image.split(",")[1];
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", (payload.name || "user") + "_" + Date.now());
        var file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoId = file.getId();
      } catch (err) {
        console.error("Photo Error: " + err.toString());
      }
    }
    
    sheet.appendRow([
      now, 
      Utilities.formatDate(now, "GMT+9", "HH:mm:ss"),
      payload.name,
      payload.phone,
      payload.type,
      payload.item,
      payload.comment || "",
      photoId,
      payload.score || 0
    ]);
    
    return { success: true, photoId: photoId };
  } catch (e) {
    return { success: false, error: "기록 저장 실패: " + e.toString() };
  }
}

/**
 * [대시보드] 특정 회원의 실시간 데이터(정보 + 점수) 가져오기 (v44.146)
 */
function getUserDashboardData(payload) {
  try {
    var phone = String(payload.phone || "").replace(/[^0-9]/g, ""); // [v44.173] 숫자만 남기도록 정규화
    if (!phone) return { error: "전화번호가 없습니다." };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 회원 정보 (등록현황 시트)
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    var memberInfo = { name: "모험가", tier: "씨앗", rank: "-" };
    if (regSheet) {
      var regData = regSheet.getDataRange().getDisplayValues();
      var regCols = getRegColumnIndices(regSheet);
      for (var i = 1; i < regData.length; i++) {
        var sheetPhone = String(regData[i][regCols.phone]).replace(/[^0-9]/g, ""); 
        if (sheetPhone === phone || sheetPhone.indexOf(phone) > -1 || phone.indexOf(sheetPhone) > -1) {
          memberInfo.name = regData[i][regCols.name];
          memberInfo.tier = regData[i][regCols.membership] || "새싹";
          break;
        }
      }
    }

    // [v44.165] 일일 출석 점수(5점) 자동 지급 로직
    var arcSheet = ss.getSheetByName("아카이브") || ss.insertSheet("아카이브");
    var nowRef = new Date(); // 기준 시각
    var todayStr = Utilities.formatDate(nowRef, "GMT+9", "yyyy-MM-dd");
    var isFirstLoginToday = false;
    
    // [v44.170] 중복 기록 방지를 위한 서버 잠금(Lock) 도입
    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000); 
      // [v44.174] 32분의 저주 방지를 위해 getDisplayValues 사용
      var arcDataCheck = arcSheet.getDataRange().getDisplayValues(); 
      var alreadyLogged = false;
      for (var j = 1; j < arcDataCheck.length; j++) {
        // [v44.175] 날짜 형식 정규화 (2026. 5. 15 -> 2026-05-15)
        var rawDateStr = arcDataCheck[j][0].split(" ")[0].replace(/\./g, "-").replace(/-$/, "");
        var parts = rawDateStr.split("-").filter(function(p){ return p !== ""; });
        var normDateStr = "";
        if (parts.length >= 3) {
          var y = parts[0];
          var m = parts[1].length === 1 ? "0" + parts[1] : parts[1];
          var d = parts[2].length === 1 ? "0" + parts[2] : parts[2];
          normDateStr = y + "-" + m + "-" + d;
        }

        var recPhone = String(arcDataCheck[j][3]).replace(/[^0-9]/g, "");
        var recType = String(arcDataCheck[j][4]);
        if (normDateStr === todayStr && recPhone === phone && recType === "로그인") {
          alreadyLogged = true;
          break;
        }
      }
      
      if (!alreadyLogged) {
        arcSheet.appendRow([
          nowRef,
          Utilities.formatDate(nowRef, "GMT+9", "HH:mm:ss"),
          memberInfo.name,
          phone,
          "로그인",
          "일일 출석",
          "지니 월드 입장 완료",
          "", 
          5   
        ]);
        isFirstLoginToday = true;
      }
    } catch (e) {
      console.error("Lock error: " + e.toString());
    } finally {
      lock.releaseLock();
    }

    // 2. 능력치 및 점수 계산 (아카이브 시트)
    var stats = { health: 0, perf: 0, def: 0 };
    var scores = { lifetime: 0, season: 0, weekly: 0 };
    
    var calcNow = new Date();
    // 주간 기준 (이번 주 월요일 00:00)
    var startOfWeek = new Date(calcNow.getFullYear(), calcNow.getMonth(), calcNow.getDate() - calcNow.getDay() + (calcNow.getDay() === 0 ? -6 : 1));
    startOfWeek.setHours(0, 0, 0, 0);
    // 시즌 기준 (v39: 현재 월의 1일)
    var startOfSeason = new Date(calcNow.getFullYear(), calcNow.getMonth(), 1);
    startOfSeason.setHours(0, 0, 0, 0);

    if (arcSheet) {
      // [v44.174] 32분의 저주 방지를 위해 getDisplayValues 사용
      var arcData = arcSheet.getDataRange().getDisplayValues();
      for (var j = 1; j < arcData.length; j++) {
        var arcPhone = String(arcData[j][3]).replace(/[^0-9]/g, ""); 
        if (arcPhone === phone) {
          var score = Number(arcData[j][8] || 0);
          // 문자열 날짜를 안전하게 Date 객체로 변환 (시간은 00:00:00 고정)
          var dateParts = arcData[j][0].split("-");
          var recDate = new Date(dateParts[0], dateParts[1] - 1, dateParts[2]);
          
          // 1) 누적 (Lifetime)
          scores.lifetime += score;
          // 2) 시즌 (Season)
          if (recDate >= startOfSeason) scores.season += score;
          // 3) 주간 (Weekly)
          if (recDate >= startOfWeek) {
            scores.weekly += score;
            // 능력치 분배 (주간 뷰용)
            var type = String(arcData[j][4] || "");
            if (type.indexOf("퀘스트") > -1) stats.perf += score;
            else if (type.indexOf("습관") > -1) stats.def += score;
            else if (type.indexOf("인바디") > -1 || type.indexOf("건강") > -1) stats.health += score;
            else { 
              stats.health += (score * 0.3);
              stats.perf += (score * 0.4);
              stats.def += (score * 0.3);
            }
          }
        }
      }
    }

    // 3. 인바디 점수 합산 (33챌린지_인바디)
    var inbodySheet = ss.getSheetByName("33챌린지_인바디");
    if (inbodySheet) {
      var inbodyData = inbodySheet.getDataRange().getValues();
      for (var k = 1; k < inbodyData.length; k++) {
        var inbodyPhone = String(inbodyData[k][2]).replace(/[^0-9]/g, "");
        if (inbodyPhone === phone) {
          var inbodyScore = Number(inbodyData[k][6] || 0);
          var inDate = new Date(inbodyData[k][0]);
          scores.lifetime += inbodyScore;
          if (inDate >= startOfSeason) scores.season += inbodyScore;
          if (inDate >= startOfWeek) {
            scores.weekly += inbodyScore;
            stats.health += inbodyScore;
          }
        }
      }
    }

    // 4. v39 7단계 티어 및 진화 정보 산출
    var tiers = [
      { name: "씨앗 🌱", min: 0, max: 1000 },
      { name: "새싹 🌿", min: 1001, max: 3000 },
      { name: "나무 🌳", min: 3001, max: 8000 },
      { name: "꽃 🌸", min: 8001, max: 15000 },
      { name: "꿈나무 요정 🧚‍♂️", min: 15001, max: 30000 },
      { name: "전설의 점퍼 👑", min: 3001, max: 60000 },
      { name: "지니 월드 수호신 🌌", min: 60001, max: 9999999 }
    ];
    
    var currentTier = tiers[0];
    var nextTier = tiers[1];
    for (var t = 0; t < tiers.length; t++) {
      if (scores.lifetime >= tiers[t].min) {
        currentTier = tiers[t];
        nextTier = tiers[t+1] || tiers[t];
      } else break;
    }
    
    // 진화도 계산 (%)
    var tierRange = nextTier.max - currentTier.min;
    var progressInTier = scores.lifetime - currentTier.min;
    var evolutionPercent = Math.min(100, Math.floor((progressInTier / tierRange) * 100));

    // 5. 랭킹 산출 (시즌 점수 기준 - v39)
    var rank = "-";
    try {
      // 모든 모험가의 시즌 점수를 합산하여 순위 매기기
      var allSeasonScores = {};
      if (arcSheet) {
        var allArc = arcSheet.getDataRange().getValues();
        for (var a = 1; a < allArc.length; a++) {
          var aDate = new Date(allArc[a][0]);
          if (aDate >= startOfSeason) {
            var aPhone = String(allArc[a][3]).replace(/[^0-9]/g, "");
            var aScore = Number(allArc[a][8] || 0);
            allSeasonScores[aPhone] = (allSeasonScores[aPhone] || 0) + aScore;
          }
        }
      }
      var sortedScores = Object.values(allSeasonScores).sort(function(a, b) { return b - a; });
      var mySeasonScore = scores.season;
      rank = sortedScores.indexOf(mySeasonScore) + 1;
    } catch(e) { rank = "-"; }

    return {
      success: true,
      name: memberInfo.name,
      tier: currentTier.name,
      nextTier: nextTier.name,
      evolution: evolutionPercent,
      totalScore: scores.lifetime, 
      seasonScore: scores.season,
      weeklyScore: scores.weekly,
      rank: rank,
      isFirstLoginToday: isFirstLoginToday,
      stats: {
        weekly: stats,
        monthly: { health: stats.health * 4, perf: stats.perf * 4, def: stats.def * 4 }, // 시즌 누적 추산
        targets: { 
          total: { weekly: 1000, monthly: 4000 },
          items: {
            weekly: { health: 300, perf: 400, def: 300 }, // v39 3:4:3 비율 적용
            monthly: { health: 1200, perf: 1600, def: 1200 }
          }
        }
      }
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * [v44.160] 인바디 기록 저장 및 v39 자동 점수 계산
 */
function submitInBodyRecord(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지_인바디") || ss.insertSheet("33챌린지_인바디");
    
    var name = payload.name;
    var phone = String(payload.phone || "").replace(/[^0-9]/g, "");
    var weight = Number(payload.weight);
    var muscle = Number(payload.muscle);
    var fat = Number(payload.fat);
    
    // 이전 기록 찾기 (점수 계산용)
    var data = sheet.getDataRange().getValues();
    var prevRecord = null;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][2]).replace(/[^0-9]/g, "") === phone) {
        prevRecord = { weight: data[i][3], muscle: data[i][4], fat: data[i][5] };
        break;
      }
    }
    
    var changeScore = 0;
    if (prevRecord) {
      // v39 공식 적용 (100g = 0.1kg)
      var diffWeight = Number((prevRecord.weight - weight).toFixed(2)); // 감량(+)
      var diffMuscle = Number((muscle - prevRecord.muscle).toFixed(2)); // 증량(+)
      var diffFat = Number((prevRecord.fat - fat).toFixed(1)); // 감량(+)
      
      if (diffWeight > 0) changeScore += (diffWeight * 10) * 50; 
      if (diffMuscle > 0) changeScore += (diffMuscle * 10) * 100;
      if (diffFat > 0) changeScore += (diffFat * 10) * 75;
      
      // 유지 보너스 (+100)
      if (Math.abs(diffWeight) <= 0.2) changeScore += 100;
    }
    
    sheet.appendRow([
      new Date(),
      name,
      phone,
      weight,
      muscle,
      fat,
      changeScore,
      payload.remarks || ""
    ]);
    
    return { success: true, score: changeScore };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * [입장] 전화번호 뒷자리로 회원 검색 (키오스크 스타일 v44.148)
 */
function searchMembersByDigits(payload) {
  try {
    var digits = String(payload.digits || "").trim();
    if (digits.length < 2) return { success: true, members: [] };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록현황") || ss.getSheetByName("등록 현황");
    if (!regSheet) return { success: false, error: "'등록현황' 시트를 찾을 수 없습니다." };
    
    // getDisplayValues()를 사용하여 날짜/숫자 변환 오류 방지 (v44.164 중복 제거 적용)
    var data = regSheet.getDataRange().getDisplayValues();
    var cols = getRegColumnIndices(regSheet);
    var memberMap = {}; // [v44.164] 중복 방지를 위한 맵
    
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][cols.name] || "모험가").trim(); // 이름
      var phoneRaw = String(data[i][cols.phone] || "").trim(); // 휴대폰
      var phoneOnlyDigits = phoneRaw.replace(/[^0-9]/g, "");
      var status = String(data[i][cols.status] || "").trim(); // 상태
      
      // [출석 앱 공식] 진행중인 회원 + 뒷 4자리 엄격 매칭
      if ((status === "진행중" || status === "진행 중") && phoneOnlyDigits.slice(-4) === digits) {
        // 이미 맵에 등록된 번호라면 건너뜁니다 (중복 제거)
        if (!memberMap[phoneOnlyDigits]) {
          var phoneHint = phoneRaw.length > 4 ? "****" + phoneRaw.slice(-4) : phoneRaw;
          memberMap[phoneOnlyDigits] = { 
            name: name, 
            phone: phoneOnlyDigits,
            hint: phoneHint
          };
        }
      }
    }
    
    // 맵의 값들만 배열로 변환하여 반환
    var matches = Object.values(memberMap);
    return { success: true, members: matches.slice(0, 20) };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}


/**
 * [Vercel 지원] 외부 사이트(Vercel 등)에서 데이터를 주고받기 위한 API 핸들러
 */
function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  // 1. 웹 앱 화면을 띄워주는 경우 (구글 내부에서 직접 실행 시)
  if (!e.parameter.action) {
    var page = e.parameter.page || 'index';
    var template;
    try {
      template = HtmlService.createTemplateFromFile(page);
    } catch(err) {
      // 페이지가 없을 경우 기본 출석 페이지로
      template = HtmlService.createTemplateFromFile('attendance');
    }
    
    var title = "노형점핑 ERP";
    if (page === 'index') title = "🏰 GENIE WORLD";
    else if (page === 'reservation') title = "🔮 GENIE GATE (예약 포탈)";
    else if (page === 'challenge') title = "🏛️ GENIE GUILD (33 챌린지)";
    else if (page === 'halloffame') title = "TOP HEROES (시상자&출석왕)";
    else if (page === 'miracle') title = "MIRACLE (오운완&식단)";
    else if (page === 'notice') title = "INSIDE (건강 꿀팁 & 공지)";
    else if (page === 'community') title = "CONNECT (칭찬&수다)";
    else if (page === 'renewal') title = "💳 [연장 화면]";
    else if (page === 'challenge') title = "🏰 NOHYUNG WORLD (33 챌린지)";
    else if (page === 'admin') title = "⚙️ [관리자 화면]";
    else if (page === 'attendance') title = "🏃 [출석 화면]";
    
    template.scriptUrl = getScriptUrl();
    template.targetPhone = (e.parameter.phone || '').trim();
    template.targetName = (e.parameter.name || '').trim();
    
    return template.evaluate()
        .setTitle(title)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  
  // 2. API 요청인 경우 (Vercel 등 외부 연동 시)
  return handleRequest(e);
}

/**
 * [공용] 모든 API 요청을 처리하는 핵심 라우터
 */
function handleRequest(e) {
  var action = e.parameter.action;
  var payload = {};
  
  try {
    if (e.postData && e.postData.contents) {
      var contents = e.postData.contents;
      payload = (typeof contents === 'string') ? JSON.parse(contents) : contents;
    } else {
      payload = e.parameter || {};
    }
  } catch(err) {
    payload = e.parameter || {};
  }
  
  var result;
  try {
    // [보안강화] 전역 함수 호출 방식 개선
    var func = this[action] || globalThis[action];
    
    if (typeof func === 'function') {
      var args = (payload && payload.args) ? payload.args : [payload];
      result = func.apply(this, args);
    } else {
      result = { error: '찾을 수 없는 요청(Action)입니다: ' + action, debug: 'Function type: ' + (typeof func) };
    }
  } catch (err) {
    result = { 
      error: '서버 실행 오류', 
      message: err.toString(),
      action: action,
      stack: err.stack
    };
  }
  
  // 결과가 undefined일 경우를 위한 안전 장치
  if (result === undefined) result = { success: true, message: '작업이 완료되었으나 반환 값이 없습니다.' };
  
  return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
}






/**
 * [지능형] 시트 제목을 읽어 열 번호를 자동으로 찾아주는 함수
 * 시트 구조가 바뀌어도 에러 없이 작동하게 합니다.
 */
function getRegColumnIndices(sheet) {
  var fullData = sheet.getRange(1, 1, 5, Math.min(sheet.getLastColumn(), 20)).getValues();
  var map = { name: 1, phone: 2, status: 8, expire: 6, membership: 4, remain: 7, sign: 11, regDate: 0, headerRow: 0 }; 
  
  for (var r = 0; r < fullData.length; r++) {
    var hasName = false;
    for (var c = 0; c < fullData[r].length; c++) {
      var title = String(fullData[r][c] || "").trim().replace(/\s/g, "");
      if (title.indexOf("이름") !== -1 || title.indexOf("회원명") !== -1) {
        map.name = c;
        hasName = true;
      }
      if (title.indexOf("휴대폰") !== -1 || title.indexOf("전화번호") !== -1) map.phone = c;
      if (title.indexOf("상태") !== -1) map.status = c;
      if (title.indexOf("만료일") !== -1 || title.indexOf("종료일") !== -1) map.expire = c;
      if (title.indexOf("권종") !== -1 || title.indexOf("회원권") !== -1) map.membership = c;
      if (title.indexOf("잔여횟수") !== -1 || title.indexOf("잔여") !== -1) map.remain = c;
      if (title.indexOf("등록일") !== -1) map.regDate = c;
    }
    if (hasName) {
      map.headerRow = r;
      break; 
    }
  }
  return map;
}

function getAttendanceColumnIndices(sheet) {
  var fullData = sheet.getRange(1, 1, 5, Math.min(sheet.getLastColumn(), 20)).getValues();
  var map = { 
    date: 0, inTime: 1, name: 2, phone: 3, type: 4, 
    prev: 5, change: 6, remain: 7, reason: 8, 
    classes: 9, workoutTime: 10, status: 11, outTime: 12, memo: 13 
  }; 
  
  for (var r = 0; r < fullData.length; r++) {
    var hasKey = false;
    for (var c = 0; c < fullData[r].length; c++) {
      var title = String(fullData[r][c] || "").trim().replace(/\s/g, "");
      if (title.indexOf("날짜") !== -1) map.date = c;
      else if (title.indexOf("기록시간") !== -1 || title.indexOf("입실시간") !== -1) map.inTime = c;
      else if (title.indexOf("이름") !== -1) { map.name = c; hasKey = true; }
      else if (title.indexOf("휴대폰") !== -1) map.phone = c;
      else if (title.indexOf("유형") !== -1) map.type = c;
      else if (title.indexOf("변동") !== -1 || title.indexOf("차감") !== -1) map.change = c;
      else if (title.indexOf("잔여") !== -1) map.remain = c;
      else if (title.indexOf("사유") !== -1) map.reason = c;
      else if (title.indexOf("참여클래스") !== -1 || title.indexOf("클래스") !== -1) map.classes = c;
      else if (title.indexOf("상태") !== -1) map.status = c;
      else if (title.indexOf("퇴실시간") !== -1) map.outTime = c;
      else if (title.indexOf("비고") !== -1) map.memo = c;
    }
    if (hasKey) break;
  }
  return map;
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl().split('?')[0];
}

/**
 * [원장님 전용] 시트 구조 최적화 및 강제 업데이트
 * 모든 시트 이름을 통일하고, 부족한 시트를 생성하며, 제목줄을 원장님 요청 구조로 맞춥니다.
 */
function forceUpdateAllHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- [1] 기본 운영 시트 정돈 ---
  
  // 1. 등록 현황 (띄어쓰기 포함 이름 통일)
  var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록 현황");
  if (regSheet) regSheet.setName("등록 현황");
  else regSheet = ss.insertSheet("등록 현황");
  
  // 2. 회원명단 (기존 회원DB 또는 회원명단 통일)
  var memberSheet = ss.getSheetByName("회원명단") || ss.getSheetByName("회원명단");
  if (memberSheet) memberSheet.setName("회원명단");
  else memberSheet = ss.insertSheet("회원명단");
  
  // 회원명단 헤더 (Q~V열 통계 컬럼 포함)
  var mHeaders = ["회원ID", "이름", "전화번호", "생년월일", "주소", "회원권종류", "시작일", "만료일", "잔여횟수", "보너스횟수", "등록이력 요약", "추천인", "추천인전화번호", "운동목표", "목표체중", "특이사항", "주간출석", "주간타임", "주간시간", "월간출석", "월간타임", "월간시간"];
  memberSheet.getRange(1, 1, 1, mHeaders.length).setValues([mHeaders]);

  // 3. 업무일지
  var workLogSheet = ss.getSheetByName("업무일지") || ss.insertSheet("업무일지");
  var wlHeaders = ["날짜", "작성자", "점핑목록", "근력목록", "09시", "10시", "17시", "18시", "19시", "20시", "총출석", "점핑인원", "테라피인원", "특이사항", "시설파손"];
  workLogSheet.getRange(1, 1, 1, wlHeaders.length).setValues([wlHeaders]);

  // 4. 출석기록 (기존 출석기록 + 출석_차감로그 통합)
  var logSheet = ss.getSheetByName("출석기록") || ss.getSheetByName("출석기록") || ss.insertSheet("출석기록");
  logSheet.setName("출석기록");
  var logHeaders = ["날짜", "기록시간(입실)", "이름", "회원ID", "유형", "변경전횟수", "변동치", "잔여횟수", "사유", "참여클래스", "운동타임수", "상태", "퇴실시간", "비고"];
  logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);

  // --- [2] 33 챌린지 관련 시트 정돈 (원장님 기존 양식 100% 반영) ---

  // 1. 인바디 입력
  var ibInput = ss.getSheetByName("인바디 입력") || ss.insertSheet("인바디 입력");
  var ibHeaders = ["ID", "측정일", "이름", "회원ID", "체중", "골격근량", "체지방량", "상담메모"];
  ibInput.getRange(1, 1, 1, ibHeaders.length).setValues([ibHeaders]);

  // 2. 33챌린지 주별 누적
  var ibWeekly = ss.getSheetByName("33챌린지 주별 누적") || ss.insertSheet("33챌린지 주별 누적");
  var wkHeaders = ["정산기준", "회원ID", "이름", "시즌명", "체중", "골격근량", "체지방량", "체지방률", "주간체중변화", "주간체지방변화", "주간체지방변화율", "시즌체지방변화율", "상담메모", "비고", "주간 체지방 변화율 순위", "주간 체중 감량량 순위", "시즌 누적 변화율 순위"];
  ibWeekly.getRange(1, 1, 1, wkHeaders.length).setValues([wkHeaders]);

  // 3. 33챌린지 랭킹뷰
  var ibRank = ss.getSheetByName("33챌린지 랭킹뷰") || ss.insertSheet("33챌린지 랭킹뷰");
  var rHeaders = ["정산기준", "회원ID", "이름", "시즌명", "체중", "골격근", "체지방", "체지방률", "주간체중변화", "주간체지방변화", "주간체지방변화율", "시즌체지방변화율", "상담메모"];
  ibRank.getRange(1, 1, 1, rHeaders.length).setValues([rHeaders]);

  // 4. 인바디 시즌제 등록
  var ibSeason = ss.getSheetByName("인바디 시즌제 등록") || ss.insertSheet("인바디 시즌제 등록");
  var sHeaders = ["시즌명", "시상일"];
  ibSeason.getRange(1, 1, 1, sHeaders.length).setValues([sHeaders]);

  // --- [3] 테라피 및 매출 관련 시트 ---

  // 1. 예약DB (테라피 예약)
  var resSheet = ss.getSheetByName("예약DB") || ss.insertSheet("예약DB");
  var resHeaders = ["예약ID", "회원ID", "이름", "예약날짜", "입실시간", "찜질방시작", "찜질방종료", "배정방", "인원수", "상태", "테라피시작시간", "테라피완료시간", "노쇼차감여부"];
  resSheet.getRange(1, 1, 1, resHeaders.length).setValues([resHeaders]);

  // 2. 설정 (가격 및 차감 룰)
  var configSheet = ss.getSheetByName("설정") || ss.insertSheet("설정");
  var configHeaders = ["회원권종류", "기본횟수", "유효기간(일)", "점핑차감", "테라피차감", "복합차감", "기본가격(원)"];
  configSheet.getRange(1, 1, 1, configHeaders.length).setValues([configHeaders]);
  // 기본 샘플 데이터 (시트가 비었을 때만)
  if (configSheet.getLastRow() < 2) {
    configSheet.appendRow(["점핑 30회", 30, 90, 1, 2, 3, 300000]);
    configSheet.appendRow(["테라피 10회", 10, 60, "불가", 1, "불가", 250000]);
  }

  // 3. 판매내역 (매출 장부)
  var salesSheet = ss.getSheetByName("판매내역") || ss.insertSheet("판매내역");
  var salesHeaders = ["ID", "날짜", "구분", "구입자", "항목명", "금액", "결제수단", "비고"];
  salesSheet.getRange(1, 1, 1, salesHeaders.length).setValues([salesHeaders]);

  // 4. 벙개테라피 및 휴일 설정 (명칭 통일)
  var holidaySheet = ss.getSheetByName("벙개테라피 및 휴일 설정") || ss.insertSheet("벙개테라피 및 휴일 설정");
  var hHeaders = ["날짜", "구분", "메모(벙개시간)"];
  holidaySheet.getRange(1, 1, 1, hHeaders.length).setValues([hHeaders]);

  // --- [4] 통계 및 기타 도구 시트 ---
  if (!ss.getSheetByName("주간 통계")) ss.insertSheet("주간 통계");
  if (!ss.getSheetByName("월간 통계")) ss.insertSheet("월간 통계");
  if (!ss.getSheetByName("잔여횟수 차감현황")) ss.insertSheet("잔여횟수 차감현황");
  if (!ss.getSheetByName("공지사항")) {
    var noticeSheet = ss.insertSheet("공지사항");
    noticeSheet.appendRow(["번호", "구분", "공지내용"]);
    noticeSheet.appendRow([1, "일반", "환영합니다! 노형점핑&체온테라피입니다."]);
  }

  // 5. 아카이브 (실시간 피드용)
  var archiveSheet = ss.getSheetByName("아카이브") || ss.insertSheet("아카이브");
  var arHeaders = ["날짜", "시간", "이름", "전화번호", "유형", "항목", "코멘트", "사진ID", "점수"];
  archiveSheet.getRange(1, 1, 1, arHeaders.length).setValues([arHeaders]);

  // 서식 정리
  var allSheets = [regSheet, memberSheet, workLogSheet, logSheet, ibInput, ibWeekly, ibRank, ibSeason, resSheet, configSheet, salesSheet, holidaySheet, archiveSheet];
  allSheets.forEach(function(s) {
    if (s) {
      var lastCol = s.getLastColumn();
      if (lastCol > 0) {
        s.getRange(1, 1, 1, lastCol).setFontWeight("bold").setBackground("#f1f3f5").setHorizontalAlignment("center");
        s.setFrozenRows(1);
      }
    }
  });

  SpreadsheetApp.getUi().alert("✅ [아카이브, 예약DB, 설정, 판매내역]을 포함한 모든 시트 구조가 완벽하게 정돈되었습니다!\n\n이제 ERP와 실시간 소셜 피드를 사용하실 수 있습니다.");
}



/**
 * [공용 도구] 폴더 가져오기 또는 생성
 */

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

/**
 * [자동화] 만료일이 지난 '진행중' 내역을 찾아 자동으로 '마감' 처리합니다.
 * 매일 새벽에 트리거로 실행하면 좋습니다.
 */
function autoExpireMemberships() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황");
    if (!regSheet) return;
    
    var data = regSheet.getDataRange().getValues();
    var cols = getRegColumnIndices(regSheet);
    var now = new Date();
    now.setHours(0, 0, 0, 0); 
    
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      var status = String(data[i][cols.status]).trim(); 
      var expireDateRaw = data[i][cols.expire];         
      
      if (status === "진행중" && expireDateRaw) {
        var expDate = new Date(expireDateRaw);
        expDate.setHours(0, 0, 0, 0);
        
        if (expDate < now) {
          var remainVal = data[i][cols.remain]; 
          var expireInfo = "[기한마감] 잔여 " + remainVal + "회 소멸됨 (" + Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd") + ")";
          
          regSheet.getRange(i + 1, cols.status + 1).setValue("마감(기간만료)"); 
          regSheet.getRange(i + 1, 11).setValue(expireInfo); 
          count++;
        }
      }
    }
    console.log("총 " + count + "건의 만료 내역을 마감 처리했습니다.");
  } catch (e) {
    console.error("자동 만료 처리 중 오류: " + e.toString());
  }
}

/**
 * [마이그레이션] 등록현황 시트를 15개 컬럼 -> 12개 컬럼으로 데이터 유실 없이 재배치합니다.
 */
function migrateRegistrationSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("등록 현황");
  if (!sheet) return;

  var fullData = sheet.getDataRange().getValues();
  if (fullData.length < 1) return;

  // 1. 새로운 헤더 정의 (12개)
  var newHeaders = ["등록일", "이름", "휴대폰", "결제구분", "권종", "시작일", "만료일", "잔여횟수", "상태", "결제금액", "비고", "서명"];
  var newData = [newHeaders];

  // 2. 기존 데이터가 있는 경우만 이사 시작
  for (var i = 1; i < fullData.length; i++) {
    var row = fullData[i];
    
    // 기존 구조 (15개): 등록ID(0), 등록일(1), 회원ID(2), 이름(3), 휴대폰(4), 결제구분(5), 권종(6), 시작일(7), 만료일(8), 잔여횟수(9), 상태(10), 결제금액(11), 비고(12), 기록완료(13), 서명(14)
    // 새로운 구조 (12개)에 맞춰 매칭
    if (row.length >= 12) {
      var newRow = [
        row[1],  // 등록일
        row[3],  // 이름
        row[4],  // 휴대폰
        row[5],  // 결제구분
        row[6],  // 권종
        row[7],  // 시작일
        row[8],  // 만료일
        row[9],  // 잔여횟수
        row[10], // 상태
        row[11], // 결제금액
        row[12], // 비고
        row[14]  // 서명 (기록완료 13번은 버림)
      ];
      newData.push(newRow);
    }
  }

  // 3. 시트 싹 비우고 새로 쓰기
  sheet.clear();
  sheet.getRange(1, 1, newData.length, newHeaders.length).setValues(newData);
  
  // 스타일 정리
  sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight("bold").setBackground("#d4edda");
  sheet.setFrozenRows(1);
}


// ──────────────────────────────────────────────
// 1. 초기 데이터베이스 세팅 (최초 1회만 실행)
// ──────────────────────────────────────────────
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1) 회원DB 탭 생성 (보너스 횟수 추가)
  var memberSheet = ss.getSheetByName("회원명단");
  if (!memberSheet) {
    memberSheet = ss.insertSheet("회원명단");
    memberSheet.appendRow(["회원ID", "이름", "휴대폰", "생년월일", "주소", "권종", "시작일", "만료일", "잔여횟수", "보너스횟수(추천등)", "상태", "추천인", "추천인ID", "복용약", "목표", "목표체중", "노쇼횟수", "약정서동의", "서명(이미지)", "메모"]);
    memberSheet.getRange("A1:T1").setFontWeight("bold").setBackground("#f8f9fa");
    memberSheet.setFrozenRows(1);
  }
  
  // 2) 등록현황 탭 생성 (결제 장부 - 15컬럼 정석 구조)
  var regSheet = ss.getSheetByName("등록 현황");
  if (!regSheet) {
    regSheet = ss.insertSheet("등록 현황");
    var regHeaders = ["등록ID", "등록일", "회원ID", "이름", "휴대폰", "결제구분", "권종", "시작일", "만료일", "잔여횟수", "상태", "결제금액", "비고", "기록완료", "서명"];
    regSheet.appendRow(regHeaders);
    regSheet.getRange(1, 1, 1, regHeaders.length).setFontWeight("bold").setBackground("#d4edda");
    regSheet.setFrozenRows(1);
  }
  
  // 3) 출석기록 탭 생성 (통합 구조)
  var logSheet = ss.getSheetByName("출석기록");
  if (!logSheet) {
    logSheet = ss.insertSheet("출석기록");
    var logHeaders = ["날짜", "기록시간(입실)", "이름", "회원ID", "유형", "변경전횟수", "변동치", "잔여횟수", "사유", "참여클래스", "운동타임수", "상태", "퇴실시간", "비고"];
    logSheet.appendRow(logHeaders);
    logSheet.getRange(1, 1, 1, logHeaders.length).setFontWeight("bold").setBackground("#fff3cd");
    logSheet.setFrozenRows(1);
  }
  
  // 4) 설정 탭 생성 (회원권 차감 룰 및 가격)
  var configSheet = ss.getSheetByName("설정");
  if (!configSheet) {
    configSheet = ss.insertSheet("설정");
    configSheet.appendRow(["회원권종류", "기본횟수", "유효기간(개월)", "점핑차감", "테라피차감", "복합차감", "기본가격(원)"]);
    configSheet.appendRow(["점핑 30회", 30, 3, 1, 2, 3, 300000]);
    configSheet.appendRow(["테라피 10회", 10, 2, "불가", 1, "불가", 250000]);
    configSheet.appendRow(["월권", 20, 1, 1, "불가", "불가", 150000]);
    configSheet.appendRow(["운동만", 20, 1, 1, "불가", "불가", 130000]);
    configSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#e3fafd");
    configSheet.setFrozenRows(1);
  }
  
  // 5) 예약DB 탭 생성 (테라피 비서 연동용 - 13컬럼 구조)
  var reserveSheet = ss.getSheetByName("예약DB");
  if (!reserveSheet) {
    reserveSheet = ss.insertSheet("예약DB");
    reserveSheet.appendRow(["예약ID", "회원ID", "이름", "예약날짜", "입실시간", "찜질방시작", "찜질방종료", "배정방", "인원수", "상태", "테라피시작시간", "테라피완료시간", "노쇼차감여부"]);
    reserveSheet.getRange("A1:M1").setFontWeight("bold").setBackground("#cce5ff");
    reserveSheet.setFrozenRows(1);
  }
  
  // 6) 문자발송 탭 생성 (등록/연장 알림용)
  var smsSheet = ss.getSheetByName("문자발송");
  if (!smsSheet) {
    smsSheet = ss.insertSheet("문자발송");
    smsSheet.appendRow(["기록시간", "이름", "전화번호", "안내분류", "생성된문자내용", "상태"]);
    smsSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#f8d7da");
    smsSheet.setFrozenRows(1);
  }
  
  // 7) 주간통계 탭 생성
  var weeklySheet = ss.getSheetByName("주간통계");
  if (!weeklySheet) {
    weeklySheet = ss.insertSheet("주간통계");
    weeklySheet.appendRow(["기록시간", "이름", "전화번호", "해당주차", "주간출석횟수", "주간운동타임수", "발송코멘트"]);
    weeklySheet.getRange("A1:G1").setFontWeight("bold").setBackground("#e2e3e5");
    weeklySheet.setFrozenRows(1);
  }

  // 8) 월간통계 탭 생성
  var monthlySheet = ss.getSheetByName("월간통계");
  if (!monthlySheet) {
    monthlySheet = ss.insertSheet("월간통계");
    monthlySheet.appendRow(["기록시간", "이름", "전화번호", "해당월", "월간출석횟수", "월간운동타임수", "발송코멘트"]);
    monthlySheet.getRange("A1:G1").setFontWeight("bold").setBackground("#e2e3e5");
    monthlySheet.setFrozenRows(1);
  }

  // 8) 인바디_시즌등록 탭 생성 (33챌린지 용)
  var seasonSheet = ss.getSheetByName("인바디_시즌등록");
  if (!seasonSheet) {
    seasonSheet = ss.insertSheet("인바디_시즌등록");
    seasonSheet.appendRow(["시즌명", "시상일(목요일)", "3주차시작일", "2주차시작일", "1주차시작일", "시작주", "상태"]);
    seasonSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#d1ecf1");
    seasonSheet.setFrozenRows(1);
  }

  // 9) 인바디_기록 탭 생성 (33챌린지 용)
  var inbodySheet = ss.getSheetByName("인바디_기록");
  if (!inbodySheet) {
    inbodySheet = ss.insertSheet("인바디_기록");
    inbodySheet.appendRow(["기록일자", "이름", "전화번호", "시즌명", "주차구분", "체중", "체지방량", "근육량", "감량점수", "주간순위"]);
    inbodySheet.getRange("A1:J1").setFontWeight("bold").setBackground("#d1ecf1");
    inbodySheet.setFrozenRows(1);
  }

  // 10) 업무일지 탭 생성 (교대 인수인계 용)
  var workLogSheet = ss.getSheetByName("업무일지");
  if (!workLogSheet) {
    workLogSheet = ss.insertSheet("업무일지");
    workLogSheet.appendRow(["날짜", "작성자(미진/현정)", "점핑목록", "근력목록", "09시출석", "10시출석", "17시출석", "18시출석", "19시출석", "20시출석", "전체출석수", "테라피인원", "특이사항", "시설고장_파손"]);
    workLogSheet.getRange("A1:N1").setFontWeight("bold").setBackground("#fff3cd");
    workLogSheet.setFrozenRows(1);
  }

  // 11) 판매내역 탭 생성 (8컬럼 - 원장님 기존 구조)
  var salesSheet = ss.getSheetByName("판매내역");
  if (!salesSheet) {
    salesSheet = ss.insertSheet("판매내역");
    salesSheet.appendRow(["ID", "날짜", "구분", "구입자", "항목명", "금액", "결제수단", "비고"]);
    salesSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#d4edda");
    salesSheet.setFrozenRows(1);
  }



  // 13) 공지사항 탭 생성
  var noticeSheet = ss.getSheetByName("공지사항");
  if (!noticeSheet) {
    noticeSheet = ss.insertSheet("공지사항");
    noticeSheet.appendRow(["번호", "구분", "공지내용"]);
    noticeSheet.appendRow([1, "일반", "환영합니다! 노형점핑&체온테라피입니다."]);
    noticeSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#fff3cd");
    noticeSheet.setFrozenRows(1);
  }

  // 14) 벙개테라피 및 휴일 설정 탭 생성 (명칭 통일)
  var holidaySheet = ss.getSheetByName("벙개테라피 및 휴일 설정");
  if (!holidaySheet) {
    holidaySheet = ss.insertSheet("벙개테라피 및 휴일 설정");
    holidaySheet.appendRow(["날짜", "구분", "메모(벙개시간)"]);
    holidaySheet.getRange("A1:C1").setFontWeight("bold").setBackground("#f8d7da");
    holidaySheet.setFrozenRows(1);
  }

  // 기본 시트(시트1) 삭제
  var defaultSheet = ss.getSheetByName("시트1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }
  
  SpreadsheetApp.getUi().alert("✅ 14개의 마스터 DB 시트 세팅이 완벽하게 끝났습니다!");
}


// ──────────────────────────────────────────────
// 3. 백엔드 API: 회원 검색 (태블릿 출석체크용)
// ──────────────────────────────────────────────
function searchMemberByPin(pinStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    if (!regSheet) return { error: "등록현황 시트를 찾을 수 없습니다." };

    // 원장님 말씀대로 getDisplayValues()를 사용하여 날짜/숫자 변환 오류를 원천 차단
    var data = regSheet.getDataRange().getDisplayValues();
    var cols = getRegColumnIndices(regSheet);
    
    // 회원명단 시트에서 보너스 횟수 조회를 위한 준비
    var memberSheet = ss.getSheetByName("회원명단") || ss.getSheetByName("회원명단");
    var mData = memberSheet ? memberSheet.getDataRange().getDisplayValues() : [];
    
    var memberMap = {};
    
    var pin = String(pinStr || "").trim();
    if (pin.length < 4) return { error: "뒷자리 4자리를 정확히 입력해주세요." };

    for (var i = 1; i < data.length; i++) {
      var phoneRaw = data[i][cols.phone];
      var phoneClean = phoneRaw.replace(/[^0-9]/g, ""); 
      var status = String(data[i][cols.status] || "").trim(); 
      if ((status === "진행중" || status === "진행 중") && phoneClean.slice(-4) === pin) {
        if (!memberMap[phoneClean]) {
          // 회원명단에서 보너스 횟수 찾기
          var bonus = "0";
          if (mData.length > 0) {
            for (var mIdx = 1; mIdx < mData.length; mIdx++) {
              var mPhone = String(mData[mIdx][2] || "").replace(/[^0-9]/g, "");
              if (mPhone === phoneClean) {
                bonus = String(mData[mIdx][9] || "0"); // 9번 컬럼: 보너스횟수
                break;
              }
            }
          }

          memberMap[phoneClean] = {
            name: String(data[i][cols.name] || "이름없음"),
            phone: phoneRaw,
            bonusCount: bonus,
            passes: []
          };
        }
        
        memberMap[phoneClean].passes.push({
          membershipType: String(data[i][cols.membership] || "일반"),
          expireDate: data[i][cols.expire], // getDisplayValues() 덕분에 이미 예쁜 날짜 문자열임
          remainCount: String(data[i][cols.remain] || "0")
        });
      }
    }
    
    var keys = Object.keys(memberMap);
    if (keys.length === 0) {
      return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };
    }
    
    var finalResults = [];
    for (var j = 0; j < keys.length; j++) {
      var m = memberMap[keys[j]];
      finalResults.push({
        name: m.name,
        phone: m.phone,
        membershipType: m.passes.map(function(p) { return p.membershipType; }).join(" / "),
        expireDate: m.passes[0].expireDate,
        remainCount: m.passes[0].remainCount,
        bonusCount: m.bonusCount,
        allPasses: m.passes 
      });
    }
    
    // 추가: 테라피 입실 중인지 예약DB 확인 (퇴실 버튼 표시용)
    var isTherapyActive = false;
    try {
      var resSheet = ss.getSheetByName("예약DB");
      var resData = resSheet.getDataRange().getDisplayValues();
      var todayFormatted = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
      var searchPin = String(pinStr || "").trim();
      
      for (var r = 1; r < resData.length; r++) {
        var rDateStr = String(resData[r][3]).replace(/[^0-9]/g, "");
        var todayNum = todayFormatted.replace(/[^0-9]/g, "");
        var rPhone = String(resData[r][1]).replace(/[^0-9]/g, "");
        var rStatus = String(resData[r][9]);
        
        if (rDateStr === todayNum && rPhone.slice(-4) === searchPin && rStatus === "테라피 진행중") {
          isTherapyActive = true;
          break;
        }
      }
    } catch(e) {}
    
    return { success: true, members: finalResults, isTherapyActive: isTherapyActive };
  } catch (e) {
    return { error: "조회 중 서버 오류: " + e.toString() };
  }
}

// ──────────────────────────────────────────────
// 4. 백엔드 API: 출석 처리 및 차감 (태블릿 출석체크용)
// ──────────────────────────────────────────────
function processAttendance(phoneStr, type, isBonus) {
  phoneStr = formatPhoneNumber(phoneStr);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName("회원명단");
    var logSheet = ss.getSheetByName("출석기록");
    var configSheet = ss.getSheetByName("설정");
    
    // [중복 방지] 30분 이내 동일 번호 출석 여부 확인
    var now = new Date();
    var lastRow = logSheet.getLastRow();
    if (lastRow > 1) {
      // 최근 20개 기록만 확인 (성능 최적화)
      var checkRange = Math.max(2, lastRow - 20);
      var lastLogs = logSheet.getRange(checkRange, 1, (lastRow - checkRange + 1), 5).getValues();
      for (var i = lastLogs.length - 1; i >= 0; i--) {
        var lDate = lastLogs[i][0]; // A열 날짜
        var lTime = lastLogs[i][1]; // B열 시간
        var lPhone = formatPhoneNumber(String(lastLogs[i][3])); // D열 전화번호 정규화 후 비교
        
        if (lPhone === phoneStr) {
          var logDateTime = new Date(lDate);
          if (lTime instanceof Date) {
            logDateTime.setHours(lTime.getHours(), lTime.getMinutes(), lTime.getSeconds());
          } else {
            var tParts = String(lTime).split(":");
            if (tParts.length >= 2) {
              logDateTime.setHours(tParts[0], tParts[1], tParts[2] || 0);
            }
          }
          
          var diffMinutes = (now.getTime() - logDateTime.getTime()) / (1000 * 60);
          if (diffMinutes >= 0 && diffMinutes < 30) {
            return { error: Math.floor(diffMinutes) + "분 전에 이미 출석처리 되었습니다." };
          }
          break; // 가장 최근 기록만 찾으면 중단
        }
      }
    }
    
    var mData = memberSheet.getDataRange().getValues();
    var cData = configSheet.getDataRange().getValues();
    
    // 1. 모든 활성 회원권 찾기 (등록현황 시트 기준)
    var regSheet = ss.getSheetByName("등록 현황");
    var rData = regSheet.getDataRange().getValues();
    var activePasses = [];
    for (var i = 1; i < rData.length; i++) {
      if (String(rData[i][2]) === phoneStr && String(rData[i][8]).trim() === "진행중") {
        activePasses.push({
          rowIdx: i + 1,
          name: rData[i][1],
          membershipType: String(rData[i][4]),
          remainCount: rData[i][7],
          memo: rData[i][10] // K열 비고(이월딱지)
        });
      }
    }
    
    // 회원DB에서 보너스 횟수만 따로 가져오기
    var mData = memberSheet.getDataRange().getValues();
    var bonusCount = 0;
    var mRowIdx = -1;
    for (var i = 1; i < mData.length; i++) {
      if (String(mData[i][2]) === phoneStr) {
        bonusCount = mData[i][9];
        mRowIdx = i + 1;
        break;
      }
    }
    
    if (activePasses.length === 0 && bonusCount <= 0) return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };

    // 2. 설정 데이터 맵핑
    var configRules = {};
    for (var j = 1; j < cData.length; j++) {
      configRules[cData[j][0]] = { Jumping: cData[j][3], Therapy: cData[j][4], Combo: cData[j][5] };
    }

    // 3. 우선순위 로직에 따른 회원권 선택 및 차감액 결정
    var selectedPasses = []; 
    
    if (type === '점핑') {
      var pass = null;
      for (var pIdx = 0; pIdx < activePasses.length; pIdx++) {
        var p = activePasses[pIdx];
        if (p.membershipType.indexOf("점핑") !== -1 || p.membershipType.indexOf("월권") !== -1 || p.membershipType.indexOf("운동만") !== -1) {
          pass = p;
          break;
        }
      }
      if (!pass) pass = activePasses[0];
      if (!pass) return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };
      
      var rule = configRules[pass.membershipType];
      var deductAmount = 1; // 기본값 1회 차감
      
      if (rule && rule.Jumping !== "불가") {
        deductAmount = Number(rule.Jumping);
      } else if (!rule) {
        // 설정 시트에 없는 레거시(50회권 등) 처리: 이름에 '점핑'이나 '회'가 있으면 1회 차감 허용
        if (pass.membershipType.indexOf("점핑") !== -1 || pass.membershipType.indexOf("회") !== -1) {
          deductAmount = 1;
        } else {
          return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };
        }
      } else {
        return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };
      }
      
      selectedPasses.push({ pass: pass, amount: deductAmount, reason: "정상 차감 (점핑)" });
    } 
    else if (type === '테라피') {
      var therapyPass = null;
      for (var pIdx = 0; pIdx < activePasses.length; pIdx++) {
        if (activePasses[pIdx].membershipType.indexOf("테라피") !== -1) {
          therapyPass = activePasses[pIdx];
          break;
        }
      }
      
      if (therapyPass) {
        selectedPasses.push({ pass: therapyPass, amount: 1, reason: "테라피권 우선 차감" });
      } else {
        var anyJumpingSessionPass = null;
        for (var pIdx = 0; pIdx < activePasses.length; pIdx++) {
          var pType = activePasses[pIdx].membershipType;
          if (pType.indexOf("점핑") !== -1 && pType.indexOf("회") !== -1) {
            anyJumpingSessionPass = activePasses[pIdx];
            break;
          }
        }
        
        if (anyJumpingSessionPass) {
          selectedPasses.push({ pass: anyJumpingSessionPass, amount: 2, reason: "테라피권 없음 -> [" + anyJumpingSessionPass.membershipType + "]에서 2회 차감" });
        } else {
          var pass = activePasses[0];
          if (!pass) return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };
          var rule = configRules[pass.membershipType];
          if (!rule || rule.Therapy === "불가") return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };
          selectedPasses.push({ pass: pass, amount: Number(rule.Therapy), reason: "정상 차감 (테라피)" });
        }
      }
    } 
    else if (type === '복합') {
      var anyJumpingSessionPass = null;
      for (var pIdx = 0; pIdx < activePasses.length; pIdx++) {
        var pType = activePasses[pIdx].membershipType;
        if (pType.indexOf("점핑") !== -1 && pType.indexOf("회") !== -1) {
          anyJumpingSessionPass = activePasses[pIdx];
          break;
        }
      }
      
      if (anyJumpingSessionPass) {
        selectedPasses.push({ pass: anyJumpingSessionPass, amount: 3, reason: "복합출석: [" + anyJumpingSessionPass.membershipType + "]에서 3회 차감" });
      } else {
        return { error: "복합 이용은 [점핑 회수권]이 있어야 가능합니다." };
      }
    }
    else if (type === '보너스') {
      if (bonusCount <= 0) return { error: "사용 가능한 보너스 횟수가 없습니다." };
      selectedPasses.push({ isBonusType: true, amount: 1, reason: "보너스(이월) 횟수 사용" });
    }

    // 4. 실제 차감 실행
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, "GMT+9", "HH:mm:ss");
    var firstMemberName = activePasses.length > 0 ? activePasses[0].name : "회원";
    var nextCount = 0;

    // 복합 출석의 경우, 로그를 한 줄로 합쳐서 기록하기 위한 버퍼
    var comboLogInfo = { prev: 0, change: 0, remain: 0, reason: "" };

    for (var k = 0; k < selectedPasses.length; k++) {
      var item = selectedPasses[k];
      var pass = item.pass;
      var deductAmount = item.amount;
      var prevCount;
      
      if (item.isBonusType) {
        prevCount = bonusCount;
        nextCount = prevCount - deductAmount;
        memberSheet.getRange(mRowIdx, 10).setValue(nextCount);
        // [14컬럼 매칭]: 날짜, 시간, 이름, 휴대폰, 유형, 전, 변, 후, 사유, 클래스, 타임, 상태, 퇴실, 비고
        logSheet.appendRow([dateStr, timeStr, firstMemberName, phoneStr, "보너스권", prevCount, "-" + deductAmount, nextCount, item.reason, "", "", "입실", "", ""]);
      } else {
        var deductRemaining = deductAmount;
        var memoText = String(pass.memo || "");
        var carryTagRegex = /\[월권이월:(\d+)회\|기한:([0-9-]{10})\]/;
        var carryMatch = memoText.match(carryTagRegex);

        if (carryMatch && !isBonus) { 
          var carryCount = Number(carryMatch[1]);
          var carryExpDate = new Date(carryMatch[2]);
          carryExpDate.setHours(23, 59, 59, 999);
          if (now <= carryExpDate && carryCount > 0) {
            var useFromCarry = Math.min(carryCount, deductRemaining);
            var newCarry = carryCount - useFromCarry;
            var newMemo = newCarry > 0 ? memoText.replace(carryTagRegex, "[월권이월:" + newCarry + "회|기한:" + carryMatch[2] + "]") : memoText.replace(carryTagRegex, "[이월횟수 모두 사용됨]");
            regSheet.getRange(pass.rowIdx, 11).setValue(newMemo.trim()); 
            deductRemaining -= useFromCarry;
            item.reason += " (이월딱지 " + useFromCarry + "회 차감)";
          }
        }

        prevCount = pass.remainCount;
        if (deductRemaining > 0) {
          if (prevCount !== "(무제한)" && prevCount !== "무제한") {
            var curVal = Number(prevCount) || 0;
            if (curVal < deductRemaining) return { error: "[" + pass.membershipType + "] 잔여 횟수가 부족합니다." };
            nextCount = curVal - deductRemaining;
            regSheet.getRange(pass.rowIdx, 8).setValue(nextCount); 
          } else { nextCount = prevCount; }
        } else { nextCount = prevCount; }

        // 복합(Combo) 기록 합치기 로직
        if (type === '복합') {
          comboLogInfo.prev = (comboLogInfo.prev === 0) ? prevCount : comboLogInfo.prev;
          comboLogInfo.change += deductAmount;
          comboLogInfo.remain = nextCount;
          comboLogInfo.reason += (comboLogInfo.reason ? " + " : "") + item.reason;
        } else {
          // 일반 로그 기록
          logSheet.appendRow([dateStr, timeStr, firstMemberName, phoneStr, pass.membershipType, prevCount, "-" + deductAmount, nextCount, item.reason, "", "", "입실", "", ""]);
        }
      }
    }

    // 복합 로그는 루프 종료 후 한 번만 기록
    if (type === '복합' && selectedPasses.length > 0) {
      logSheet.appendRow([dateStr, timeStr, firstMemberName, phoneStr, "복합", comboLogInfo.prev, "-" + comboLogInfo.change, comboLogInfo.remain, comboLogInfo.reason, "", "", "입실", "", ""]);
      nextCount = comboLogInfo.remain; // 반환값 업데이트
    }

    // --- [예약DB 실시간 연동] 테라피/복합 출석 시 예약 상태 업데이트 ---
    if (type === '테라피' || type === '복합' || type === '보너스') {
      try {
        var resSheet = ss.getSheetByName("예약DB");
        var resData = resSheet.getDataRange().getDisplayValues();
        var todayFormatted = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
        
        var todayParts = todayFormatted.match(/\d+/g);
        var todayNum = todayParts[0] + (todayParts[1].length === 1 ? "0" : "") + todayParts[1] + (todayParts[2].length === 1 ? "0" : "") + todayParts[2];
        var phoneClean = String(phoneStr || "").replace(/[^0-9]/g, "");
        
        for (var rIdx = 1; rIdx < resData.length; rIdx++) {
          var rDateRaw = resData[rIdx][3];
          var rDateParts = String(rDateRaw).match(/\d+/g);
          if (!rDateParts || rDateParts.length < 3) continue;
          var rDateNum = rDateParts[0] + (rDateParts[1].length === 1 ? "0" : "") + rDateParts[1] + (rDateParts[2].length === 1 ? "0" : "") + rDateParts[2];
          
          var rPhone = String(resData[rIdx][1]).replace(/[^0-9]/g, "");
          var isPhoneMatch = (rPhone === phoneClean) || 
                             (rPhone.length >= 8 && phoneClean.length >= 8 && rPhone.slice(-8) === phoneClean.slice(-8));

          // 이름 매칭은 생략하고 날짜와 전화번호만으로 매칭 (이름 오타/공백 방지)
          if (rDateNum === todayNum && isPhoneMatch) {
            var status = String(resData[rIdx][9]);
            if (status.indexOf("예약") !== -1 || status.indexOf("테라피") !== -1) {
              resSheet.getRange(rIdx + 1, 10).setValue("테라피중");
              resSheet.getRange(rIdx + 1, 11).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm:ss"));
              break;
            }
          }
        }
      } catch (resErr) {
        console.error("예약DB 연동 오류: " + resErr.toString());
      }
    }

    // 실제로 사용된 회원권 이름 및 기한 찾기
    var usedPassName = "";
    var usedExpireDate = "";
    if (type === '보너스') {
      usedPassName = "보너스권";
      usedExpireDate = "-";
    } else if (type === '복합') {
      usedPassName = "복합(점핑30회)";
      // 복합은 pass 객체를 루프에서 찾음
      if (selectedPasses.length > 0 && selectedPasses[0].pass) {
        var pIdx = selectedPasses[0].pass.rowIdx;
        usedExpireDate = regSheet.getRange(pIdx, 7).getDisplayValue(); // 7번열이 기한
      }
    } else if (selectedPasses.length > 0 && selectedPasses[0].pass) {
      usedPassName = selectedPasses[0].pass.membershipType;
      var pIdx = selectedPasses[0].pass.rowIdx;
      usedExpireDate = regSheet.getRange(pIdx, 7).getDisplayValue();
    }

    return { 
      success: true, 
      message: firstMemberName + "님 출석 완료!",
      updatedRemain: nextCount,
      usedPassName: usedPassName,
      expireDate: usedExpireDate,
      isUnlimited: (nextCount === "(무제한)" || nextCount === "무제한")
    };
  } catch (e) { return { error: "서버 오류: " + e.toString() }; }
}

// ──────────────────────────────────────────────
// 5. HTML 파일 내에서 다른 파일을 포함(include) 시키는 함수
// (CSS, JS를 분리해서 관리하기 위함)
// ──────────────────────────────────────────────
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 키오스크용 퇴실 처리 (예약DB 업데이트전용)
 */
function processKioskCheckout(phoneStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resSheet = ss.getSheetByName("예약DB");
    var resData = resSheet.getDataRange().getDisplayValues();
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    var phoneClean = String(phoneStr || "").replace(/[^0-9]/g, "");
    var success = false;
    var memberName = "";

    var todayParts = todayStr.match(/\d+/g);
    var todayNum = todayParts[0] + (todayParts[1].length === 1 ? "0" : "") + todayParts[1] + (todayParts[2].length === 1 ? "0" : "") + todayParts[2];
    
    for (var i = 1; i < resData.length; i++) {
      var rDateRaw = resData[i][3];
      var rDateParts = String(rDateRaw).match(/\d+/g);
      if (!rDateParts || rDateParts.length < 3) continue;
      var rDateNum = rDateParts[0] + (rDateParts[1].length === 1 ? "0" : "") + rDateParts[1] + (rDateParts[2].length === 1 ? "0" : "") + rDateParts[2];
      
      var rPhone = String(resData[i][1]).replace(/[^0-9]/g, "");
      var isPhoneMatch = (rPhone === phoneClean) || (rPhone.length >= 8 && phoneClean.length >= 8 && rPhone.slice(-8) === phoneClean.slice(-8));
      
      if (rDateNum === todayNum && isPhoneMatch) {
        var status = String(resData[i][9]);
        if (status.indexOf("테라피") !== -1 || status.indexOf("예약") !== -1) {
          resSheet.getRange(i + 1, 10).setValue("귀가");
          resSheet.getRange(i + 1, 12).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm:ss"));
          memberName = resData[i][2];
          success = true;
          break;
        }
      }
    }

    if (success) {
      // 출석기록 시트에도 '퇴실' 상태 업데이트 (선택 사항이나 관리 일관성을 위해 추천)
      try {
        var logSheet = ss.getSheetByName("출석기록");
        var lData = logSheet.getDataRange().getValues();
        for (var j = lData.length - 1; j >= 1; j--) {
          if (String(lData[j][3]).replace(/[^0-9]/g, "") === phoneClean) {
            logSheet.getRange(j + 1, 12).setValue("귀가");
            logSheet.getRange(j + 1, 13).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm"));
            break;
          }
        }
      } catch(e) {}
      return { success: true, message: memberName + "님, 테라피가 완료되었습니다. 안녕히 가세요!" };
    } else {
      return { error: "진행 중인 테라피 내역을 찾을 수 없습니다." };
    }
  } catch(e) {
    return { error: "퇴실 처리 중 오류: " + e.toString() };
  }
}

// ──────────────────────────────────────────────
// 6. 테라피 예약 시스템 백엔드 API
// ──────────────────────────────────────────────
function getAnnouncements() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('공지사항');
    var data = sheet.getDataRange().getValues().slice(1);
    return data.map(function(row) { return row[2]; }).filter(Boolean);
  } catch(e) { return ["반갑습니다!"]; }
}

function getTodaySummary() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('예약DB');
    var flashSheet = ss.getSheetByName('벙개테라피 및 휴일 설정');
    var data = sheet.getDataRange().getDisplayValues().slice(1);
    var flashData = flashSheet.getDataRange().getDisplayValues().slice(1);
    
    var now = new Date();
    var activeDates = [];
    
    // 향후 7일 중 영업일(또는 벙개날)만 최대 5일 추출
    for (var i = 0; i < 7; i++) {
      var d = new Date();
      d.setDate(now.getDate() + i);
      var dStr = Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
      var dayOfWeek = d.getDay(); // 0:일
      
      // 휴무 설정 확인
      var isOff = (dayOfWeek === 0); // 일요일 기본 휴무
      var isFlash = false;
      
      for (var f = 0; f < flashData.length; f++) {
        var fDate = String(flashData[f][0]);
        if (fDate === dStr) {
          if (String(flashData[f][1]) === "휴무") isOff = true;
          if (String(flashData[f][2]) === "벙개") {
            isFlash = true;
            isOff = false; // 벙개면 휴무라도 영업함
          }
          break;
        }
      }
      
      if (!isOff || isFlash) {
        activeDates.push(dStr);
        if (activeDates.length >= 5) break; 
      }
    }

    var result = {};
    activeDates.forEach(date => result[date] = []);

    data.forEach(row => {
      var date = String(row[3]);
      var status = String(row[9] || "").trim();
      if (activeDates.indexOf(date) !== -1 && (status.indexOf("취소") === -1 && status !== "")) {
        var name = String(row[2]).trim();
        var maskedName = name[0] + "*" + (name.length > 2 ? name[name.length-1] : (name[1] || ""));
        
        // 시간 포맷 정규화 (H:mm -> HH:mm)
        var timeStr = String(row[4]).trim();
        var timeParts = timeStr.split(':');
        var time = "";
        if (timeParts.length >= 2) {
          var h = timeParts[0].length === 1 ? "0" + timeParts[0] : timeParts[0];
          var m = timeParts[1].length === 1 ? "0" + timeParts[1] : timeParts[1];
          time = h + ":" + m;
        } else {
          time = timeStr.substring(0, 5);
        }

        result[date].push({
          time: time,
          name: maskedName,
          room: String(row[7]),
          status: status
        });
      }
    });

    // 시간순 정렬
    activeDates.forEach(date => {
      result[date].sort((a, b) => a.time.localeCompare(b.time));
    });

    return result;
  } catch (e) { return {}; }
}

function getTodayTimetable(targetDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resSheet = ss.getSheetByName('예약DB');
    var holidaySheet = ss.getSheetByName("벙개테라피 및 휴일 설정");
    
    var targetNum = targetDate.replace(/[^0-9]/g, "");
    var dateObj = new Date(targetDate);
    var dayOfWeek = dateObj.getDay(); 

    var now = new Date();
    var todayNum = Utilities.formatDate(now, "GMT+9", "yyyyMMdd");
    var isToday = (targetNum === todayNum);

    // 기본 시간표 설정 (토요일 3타임, 평일 7타임)
    var baseTimes = (dayOfWeek === 6) ? ["09:00", "09:50", "10:40"] : ["09:00", "09:50", "10:40", "17:00", "17:50", "18:40", "19:30"];
    
    // 일요일 기본값은 휴무 (벙개 설정이 있으면 덮어씌움)
    if (dayOfWeek === 0) baseTimes = [];

    var isHoliday = false;
    var holidayMemo = "";
    var extraTimes = [];

    if (holidaySheet) {
      var holidayData = holidaySheet.getDataRange().getDisplayValues().slice(1);
      for (var i = 0; i < holidayData.length; i++) {
        var hDateNum = String(holidayData[i][0]).replace(/[^0-9]/g, "");
        if (hDateNum === targetNum) {
          var hType = String(holidayData[i][1]).trim();
          var hNote = String(holidayData[i][2]).trim();
          
          if (hType === "휴무") {
            isHoliday = true;
            holidayMemo = hNote || "센터 휴무일입니다.";
          } else if (hType === "벙개") {
            // 벙개인 경우 기존 시간 무시하고 시트의 시간만 사용하거나, 기존 시간에 추가
            // 원장님 요청: 벙개 시트에 적힌 시간대로 나오게 함
            if (hNote !== "") {
              // 콤마나 공백으로 구분된 시간을 쪼개서 추가
              var splitTimes = hNote.split(/[,/ ]+/);
              splitTimes.forEach(function(st) {
                if (st && st.indexOf(":") !== -1) extraTimes.push(st.trim());
              });
            }
          }
        }
      }
    }

    if (isHoliday) {
      return "휴무:" + (holidayMemo || "오늘은 센터 휴무일입니다.");
    }

    // 벙개 설정이 있는 경우, 벙개 시간만 보여주거나 기존 시간에 합침
    // 일요일이거나, 평일인데 특별 시간을 추가하고 싶을 때
    if (extraTimes.length > 0) {
      // 만약 일요일이라면 벙개 시간만 보여줌
      if (dayOfWeek === 0) baseTimes = extraTimes;
      else {
        // 평일이면 기존 시간에 벙개 시간을 합침 (중복 제거)
        extraTimes.forEach(function(et) {
          if (baseTimes.indexOf(et) === -1) baseTimes.push(et);
        });
      }
    } else if (dayOfWeek === 0) {
      // 일요일인데 벙개 설정조차 없으면 확실히 휴무
      return "휴무:일요일은 쉽니다.";
    }

    // [중요] 오늘 날짜인 경우, 현재 시간보다 이미 지난 타임은 필터링
    if (isToday) {
      var currentTimeStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
      baseTimes = baseTimes.filter(function(time) {
        // 예약 시간(time)이 현재 시간(currentTimeStr)보다 이후인 것만 남김
        return time >= currentTimeStr;
      });
    }

    baseTimes.sort(); // 시간순 정렬

    if (baseTimes.length === 0) {
      return "휴무:오늘은 센터 운영 시간이 없습니다.";
    }

    baseTimes.sort();
    var resData = resSheet.getDataRange().getDisplayValues().slice(1);
    var result = [];
    
    var formatToTwoDigits = function(timeStr) {
      return timeStr.split(':').map(function(v) {
        var clean = v.replace(/[^0-9]/g, "");
        return clean.length === 1 ? "0" + clean : clean;
      }).join(':').substring(0, 5);
    };

    baseTimes.forEach(function(t) {
      var targetT = formatToTwoDigits(t);

      if (isToday) {
        var tParts = targetT.split(':');
        var checkTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), Number(tParts[0]), Number(tParts[1]));
        var limitTime = new Date(now.getTime() + (20 * 60000));
        if (checkTime < limitTime) return; 
      }

      var count = resData.filter(function(row) {
        var rDateNum = row[3].replace(/[^0-9]/g, "");
        var rTimeRaw = String(row[4]).trim();
        if (!rTimeRaw || rTimeRaw === "") return false;
        var rTimeFormatted = formatToTwoDigits(rTimeRaw);
        return rDateNum === targetNum && rTimeFormatted === targetT && row[9] !== "취소";
      }).length;

      if (count >= 3) return; 

      var parts = targetT.split(':');
      var h = Number(parts[0]);
      var m = parts[1];
      var ap = (h < 12) ? "☀️ 오전 " : "🌙 오후 ";
      var dh = (h > 12) ? (h - 12) : (h === 0 ? 12 : h);
      var hs = (dh < 10) ? ("0" + dh) : dh;
      
      result.push({
        display: ap + hs + "시 " + m + "분",
        isFull: false,
        timeValue: targetT
      });
    });

    return result.length > 0 ? result : "휴무:예약 가능한 시간이 모두 마감되었습니다.";
  } catch(e) { return "에러: " + e.toString(); }
}

function getRoomStatus(date, time) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('예약DB');
  var data = sheet.getDataRange().getDisplayValues().slice(1);
  var targetDateNum = date.replace(/[^0-9]/g, ""); 
  var targetTime = time.split(':').map(function(v) { return v.length === 1 ? "0" + v : v; }).join(':').substring(0, 5);

  return ["1인실", "2인실 A", "2인실 B"].map(function(r) {
    var isBooked = data.some(function(row) {
      var rDateNum = String(row[3]).replace(/[^0-9]/g, "");
      var rTimeVal = String(row[4]).trim();
      if (!rTimeVal) return false;
      var rTimeParts = rTimeVal.split(':');
      if (rTimeParts.length < 2) return false;
      var formattedRTime = rTimeParts.slice(0, 2).map(function(v) {
        var clean = v.replace(/[^0-9]/g, "");
        return clean.length === 1 ? "0" + clean : clean;
      }).join(':');
      var rRoom = row[7];
      var rStatus = row[9];
      return rDateNum === targetDateNum && formattedRTime === targetTime && rRoom === r && rStatus !== "취소";
    });
    return { name: r, isBooked: isBooked };
  });
}

function getMyReservations(phone) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('예약DB');
  const data = sheet.getDataRange().getDisplayValues().slice(1);
  const todayNum = Number(Utilities.formatDate(new Date(), "GMT+9", "yyyyMMdd"));
  const inputPhone = String(phone).replace(/[^0-9]/g, "");

  return data.filter(function(row) {
    var sheetPhone = String(row[1]).replace(/[^0-9]/g, ""); 
    var status = String(row[9] || "").replace(/\s/g, ""); 
    
    // 날짜 정규화 (YYYYMMDD)
    var dateParts = String(row[3]).match(/\d+/g);
    var dateNum = 0;
    if (dateParts && dateParts.length >= 3) {
      var y = dateParts[0];
      var m = dateParts[1].length === 1 ? "0" + dateParts[1] : dateParts[1];
      var d = dateParts[2].length === 1 ? "0" + dateParts[2] : dateParts[2];
      dateNum = Number(y + m + d);
    }
    
    return sheetPhone === inputPhone && (status === "예약완료" || status === "테라피중") && dateNum >= todayNum;
  }).map(function(row) { 
    return { date: row[3], time: row[4], room: row[7], status: row[9] }; 
  });
}

function getMemberIDList(v) { 
  try {
    var memberSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('회원명단');
    if (!memberSheet) return [];
    var data = memberSheet.getDataRange().getDisplayValues().slice(1);
    var results = [];
    var seen = {}; // 중복 체크용

    for (var i=0; i<data.length; i++) {
      var name = String(data[i][1]).trim();
      var phoneRaw = String(data[i][2]).trim();
      var phoneOnly = phoneRaw.replace(/[^0-9]/g, "");
      var phone4 = phoneOnly.slice(-4);
      var status = String(data[i][10]).trim(); // K열 상태
      
      if (status !== "마감" && status !== "정지" && phone4 === v) {
        var key = name + "|" + phone4;
        if (!seen[key]) {
          results.push({
            displayName: name + "(" + phone4 + ")",
            name: name,
            phone: phoneRaw
          }); 
          seen[key] = true;
        }
      }
    }
    return results;
  } catch(e) { return []; }
}

function submitReservation(d) { 
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('예약DB');
    var addMinutes = function(timeStr, mins) {
      var parts = timeStr.split(':');
      var date = new Date(2000, 0, 1, parts[0], parts[1]);
      date.setMinutes(date.getMinutes() + mins);
      return Utilities.formatDate(date, "GMT+9", "HH:mm:00");
    };
    var formatTime = function(t) {
      var parts = t.split(':');
      var h = parts[0].length === 1 ? "0" + parts[0] : parts[0];
      var m = parts[1].length === 1 ? "0" + parts[1] : parts[1];
      return h + ":" + m + ":00";
    };

    var startTime = formatTime(d.time); 
    var saunaStart = addMinutes(d.time, 30);
    var saunaEnd = addMinutes(d.time, 80);
    
    var rowData = [
      Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss"), 
      d.phone, d.nameOnly, d.date, startTime, saunaStart, saunaEnd, d.room, 1, "예약 완료", "", "", ""
    ];
    sheet.appendRow(rowData);
    return "성공";
  } catch(e) { return "에러: " + e.toString(); }
}

// ──────────────────────────────────────────────
// 7. 관리자 모바일 대시보드 백엔드 API (admin.html 연동)
// ──────────────────────────────────────────────

function getAdminDashboardData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    var logSheet = ss.getSheetByName("출석기록");
    var resSheet = ss.getSheetByName("예약DB");
    
    if (!logSheet) return { error: "시트 오류: '출석기록' 시트가 없습니다. 시트 이름을 확인해주세요." };
    if (!resSheet) return { error: "시트 오류: '예약DB' 시트가 없습니다. 시트 이름을 확인해주세요." };
    
    var totalCount = 0;
    var jumpingCount = 0;
    var therapyCount = 0;
    var goHomeCount = 0;
    var activeJumping = [];
    var activeTherapy = [];
    var goHomeList = [];
    
    // 1. 출석 로그 조회 (정확한 날짜 비교를 위해 Values와 DisplayValues를 병행)
    var logData = logSheet.getDataRange().getValues();
    var logDisplayData = logSheet.getDataRange().getDisplayValues();
    var cols = getAttendanceColumnIndices(logSheet);
    
    for (var i = 1; i < logData.length; i++) {
      var dateRaw = logData[i][cols.date]; // 날짜 열 (객체 또는 문자열)
      if (!dateRaw) continue;
      
      var logDateStr = (dateRaw instanceof Date) ? Utilities.formatDate(dateRaw, "GMT+9", "yyyy-MM-dd") : String(dateRaw).split(" ")[0];
      
      if (logDateStr === todayStr) {
        var memberName = logData[i][cols.name]; 
        var memberPhone = logData[i][cols.phone];
        var membership = logData[i][cols.type]; 
        var status = String(logData[i][cols.status] || "").trim(); 
        var reason = String(logData[i][cols.reason] || "");
        
        // 시간 오차 및 형식 해결: 쌍점(:)을 기준으로 시간과 분만 추출
        var rawTimeStr = String(logDisplayData[i][cols.inTime] || "시간미정");
        var timeOnly = rawTimeStr.indexOf(":") !== -1 ? rawTimeStr.split(":").slice(0, 2).join(":") : rawTimeStr;
        
        var isCombo = reason.indexOf("복합") !== -1;
        var isTherapy = (reason.indexOf("테라피") !== -1 || reason.indexOf("보너스") !== -1) && !isCombo;
        var isJumping = (reason.indexOf("점핑") !== -1 || reason.indexOf("월권") !== -1 || reason.indexOf("운동만") !== -1) && !isTherapy && !isCombo;

        totalCount++; 

        if (isCombo) {
          jumpingCount++;
          therapyCount++;
        } else if (isTherapy) {
          therapyCount++;
        } else if (isJumping) {
          jumpingCount++;
        }

        var memberObj = {
          rowIdx: i + 1,
          name: memberName,
          phone: memberPhone,
          membership: membership,
          inTime: timeOnly,
          type: isCombo ? "복합" : (isTherapy ? "테라피" : (isJumping ? "점핑" : "보너스")),
          status: status,
          classes: String(logData[i][cols.classes] || ""), 
          timeLog: String(logData[i][cols.workoutTime] || ""), 
          extraText: String(logData[i][cols.memo] || ""), 
          reason: reason 
        };

        // 퇴실/귀가 상태 판별
        var isGoHome = (status.indexOf("귀가") !== -1 || status === "퇴실" || status === "퇴실완료");
        
        if (isGoHome) {
          goHomeCount++;
          goHomeList.push(memberObj);
        } else {
          if (isCombo || isJumping) {
            activeJumping.push(memberObj);
          } else {
            activeTherapy.push(memberObj);
          }
        }
      }
    }
    
    // 2. 테라피 예약 매칭 및 예열 알림 체크
    var resData = resSheet.getDataRange().getDisplayValues();
    var upcomingTherapy = null;
    var pendingReservations = []; // 오늘 남은 예약 명단 (입실 전)
    var todayNum = Utilities.formatDate(now, "GMT+9", "yyyyMMdd");
    
    // 예약DB 열 인덱스 더 유연하게 찾기
    function getResCols(sheet) {
      var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var map = { name: 2, date: 3, time: 4, room: 7, status: 9 };
      for (var i = 0; i < headers.length; i++) {
        var h = headers[i].toString().trim();
        if (h.indexOf("이름") !== -1) map.name = i;
        else if (h.indexOf("날짜") !== -1) map.date = i;
        else if (h.indexOf("입실시간") !== -1 || h === "시간") map.time = i;
        else if (h.indexOf("배정방") !== -1 || h.indexOf("룸") !== -1) map.room = i;
        else if (h.indexOf("상태") !== -1) map.status = i;
      }
      return map;
    }
    
    var resCols = getResCols(resSheet);
    
    for (var k = 1; k < resData.length; k++) {
      var rDateRaw = String(resData[k][resCols.date] || "").trim();
      if (!rDateRaw) continue;
      
      // 날짜 비교: 년/월/일 숫자만 뽑아서 비교 (형식 파괴 방지)
      var rDateNum = rDateRaw.replace(/[^0-9]/g, "");
      if (rDateNum.length > 8) rDateNum = rDateNum.substring(0, 8);
      // 만약 2024.5.11 처럼 월/일이 한자리인 경우 20240511 형식으로 맞춤
      if (rDateNum.length < 8) {
         var dParts = rDateRaw.split(/[^0-9]/).filter(Boolean);
         if (dParts.length >= 3) {
           rDateNum = dParts[0] + (dParts[1].length === 1 ? "0"+dParts[1] : dParts[1]) + (dParts[2].length === 1 ? "0"+dParts[2] : dParts[2]);
         }
      }
      
      if (rDateNum === todayNum) {
        var rName = String(resData[k][resCols.name] || "").trim();
        var rTimeRaw = String(resData[k][resCols.time] || "").trim();
        if (!rName || !rTimeRaw) continue;
        
        var rTimeShort = rTimeRaw.indexOf(":") !== -1 ? rTimeRaw.split(":").slice(0, 2).join(":") : rTimeRaw;
        var rRoom = String(resData[k][resCols.room] || "").trim();
        var rStatus = String(resData[k][resCols.status] || "").trim();
        
        // '예약' 글자가 포함되어 있고, '취소' 글자가 없어야 함
        if (rStatus.indexOf("예약") !== -1 && rStatus.indexOf("취소") === -1) {
          pendingReservations.push({
            name: rName,
            time: rTimeShort,
            room: rRoom,
            sortVal: rTimeShort.replace(":", "")
          });
        }
        
        // [스마트 매칭] 입실 시간과 예약 시간이 일정 범위(90분) 내인 경우에만 매칭
        var rParts = rTimeShort.split(":");
        var resHour = parseInt(rParts[0]);
        var resMinVal = parseInt(rParts[1]);
        var resTotalMin = (resHour * 60) + resMinVal;

        var nowKSTStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
        var nowParts = nowKSTStr.split(":");
        var nowTotalMin = (parseInt(nowParts[0]) * 60) + parseInt(nowParts[1]);

        for (var a=0; a<activeJumping.length; a++) {
          var mPhone = String(activeJumping[a].phone).replace(/[^0-9]/g, "");
          var rPhone = String(resData[k][1]).replace(/[^0-9]/g, "");
          var isPhoneMatch = (rPhone === mPhone) || (rPhone.length >= 8 && mPhone.length >= 8 && rPhone.slice(-8) === mPhone.slice(-8));

          if (activeJumping[a].name === rName && isPhoneMatch) {
            var inParts = activeJumping[a].inTime.split(":");
            var inTotalMin = (parseInt(inParts[0]) * 60) + parseInt(inParts[1]);
            var matchDiff = Math.abs(resTotalMin - inTotalMin);
            if (matchDiff <= 90) {
              activeJumping[a].therapyTime = rTimeShort;
              activeJumping[a].therapyRoom = rRoom;
            }
          }
        }
        for (var b=0; b<activeTherapy.length; b++) {
          var mPhoneT = String(activeTherapy[b].phone).replace(/[^0-9]/g, "");
          var rPhoneT = String(resData[k][1]).replace(/[^0-9]/g, "");
          var isPhoneMatchT = (rPhoneT === mPhoneT) || (rPhoneT.length >= 8 && mPhoneT.length >= 8 && rPhoneT.slice(-8) === mPhoneT.slice(-8));

          if (activeTherapy[b].name === rName && isPhoneMatchT) {
            var inPartsT = activeTherapy[b].inTime.split(":");
            var inTotalMinT = (parseInt(inPartsT[0]) * 60) + parseInt(inPartsT[1]);
            var matchDiffT = Math.abs(resTotalMin - inTotalMinT);
            if (matchDiffT <= 90) {
              activeTherapy[b].therapyTime = rTimeShort;
              activeTherapy[b].therapyRoom = rRoom;
            }
          }
        }
        
        if (!upcomingTherapy) {
          var diffMin = resTotalMin - nowTotalMin;
          if (diffMin > 0 && diffMin <= 30) {
            upcomingTherapy = { time: rTimeShort, name: rName, room: rRoom };
          }
        }
      }
    }
    
    // 예약 명단 시간순 정렬
    pendingReservations.sort(function(a, b) { return a.sortVal - b.sortVal; });
    
    return {
      success: true,
      totalCount: totalCount,
      jumpingCount: jumpingCount,
      therapyCount: therapyCount,
      goHomeCount: goHomeCount,
      activeJumping: activeJumping,
      activeTherapy: activeTherapy,
      goHomeList: goHomeList,
      upcomingTherapy: upcomingTherapy,
      pendingReservations: pendingReservations // 프론트엔드로 전달
    };
    
  } catch (e) { return { error: "서버 오류: " + e.toString() }; }
}

function processAdminCheckout(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    var memberSheet = ss.getSheetByName("회원명단");
    var regSheet = ss.getSheetByName("등록 현황");
    
    var cols = getAttendanceColumnIndices(logSheet);
    var rowIdx = data.rowIdx;
    var now = new Date();
    var timeStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
    
    // 1. 로그 업데이트
    logSheet.getRange(rowIdx, cols.classes + 1).setValue(data.classes); 
    logSheet.getRange(rowIdx, cols.workoutTime + 1).setValue(data.timeLog); 
    logSheet.getRange(rowIdx, cols.status + 1).setValue("귀가"); 
    logSheet.getRange(rowIdx, cols.outTime + 1).setValue(timeStr);
    
    var phoneStr = String(logSheet.getRange(rowIdx, cols.phone + 1).getValue()).replace(/[^0-9]/g, "");

    // 2. 추가 차감 처리 (체크박스 선택 시)
    if (data.extraDeduct) {
      var rData = regSheet.getDataRange().getValues();
      var rCols = getRegColumnIndices(regSheet);
      
      var targetRegIdx = -1;
      for (var i = 1; i < rData.length; i++) {
        var rPhone = String(rData[i][rCols.phone]).replace(/[^0-9]/g, "");
        if (rPhone === phoneStr && String(rData[i][rCols.status]).indexOf("진행") !== -1) {
          targetRegIdx = i + 1;
          break;
        }
      }
      
      if (targetRegIdx !== -1) {
        // 기존 사유에 추가 차감 표시
        var currentReason = String(logSheet.getRange(rowIdx, cols.reason + 1).getValue());
        if (currentReason.indexOf("(추가차감)") === -1) {
          logSheet.getRange(rowIdx, cols.reason + 1).setValue(currentReason + " (추가차감)");
        }
        logSheet.getRange(rowIdx, cols.memo + 1).setValue("기준시간 초과 (1회 추가 차감됨)");
        
        // 실제 횟수 차감 및 로그 기록 업데이트
        var remainCount = rData[targetRegIdx-1][rCols.remain];
        var isUnlimited = (remainCount === "(무제한)" || remainCount === "무제한");
        var newVal = isUnlimited ? "무제한" : (Number(remainCount) - 1);
        
        if (!isUnlimited) {
          regSheet.getRange(targetRegIdx, rCols.remain + 1).setValue(newVal);
        }
        
        // 출석기록 G열(변동치)과 H열(잔여횟수) 강제 업데이트
        var currentChange = Number(logSheet.getRange("G" + rowIdx).getValue()) || -1;
        logSheet.getRange("G" + rowIdx).setValue(currentChange - 1); 
        logSheet.getRange("H" + rowIdx).setValue(newVal);
      } else {
        logSheet.getRange(rowIdx, cols.memo + 1).setValue("⚠️회원권 못찾음 - 추가차감 실패 ★★★");
      }
    }
    
    SpreadsheetApp.flush(); 
    
    // 3. 테라피 예약 매칭 (완료 처리)
    var resSheet = ss.getSheetByName("예약DB");
    var resData = resSheet.getDataRange().getDisplayValues();
    var memberName = logSheet.getRange(rowIdx, cols.name + 1).getValue();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    var todayParts = todayStr.match(/\d+/g);
    var todayNum = todayParts[0] + (todayParts[1].length === 1 ? "0" : "") + todayParts[1] + (todayParts[2].length === 1 ? "0" : "") + todayParts[2];
    
    for (var k = 1; k < resData.length; k++) {
      var rDateRaw = resData[k][3];
      var rDateParts = String(rDateRaw).match(/\d+/g);
      if (!rDateParts || rDateParts.length < 3) continue;
      var rDateNum = rDateParts[0] + (rDateParts[1].length === 1 ? "0" : "") + rDateParts[1] + (rDateParts[2].length === 1 ? "0" : "") + rDateParts[2];
      
      var rPhone = String(resData[k][1]).replace(/[^0-9]/g, ""); 
      var isPhoneMatch = (rPhone === phoneStr) || (rPhone.length >= 8 && phoneStr.length >= 8 && rPhone.slice(-8) === phoneStr.slice(-8));

      // 이름 매칭 생략: 날짜와 전화번호만으로 매칭
      if (rDateNum === todayNum && isPhoneMatch) {
        var status = String(resData[k][9]);
        if (status.indexOf("테라피") !== -1 || status.indexOf("예약") !== -1) {
          resSheet.getRange(k + 1, 10).setValue("귀가");
          resSheet.getRange(k + 1, 12).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm:ss"));
          break;
        }
      }
    }
    
    return { success: true, message: "수정 및 퇴실 처리가 완료되었습니다." };
  } catch (e) {
    return { error: "수정 처리 오류: " + e.toString() };
  }
}

/**
 * 특정 회원의 등록현황을 모두 합산하여 회원명단의 잔여횟수를 최신화합니다.
 */


// ──────────────────────────────────────────────
// 8. 관리자 수동 출석 (이름 검색) API
// ──────────────────────────────────────────────

function searchMemberByName(nameStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var searchStr = String(nameStr || "").trim().normalize("NFC").replace(/\s/g, "").toLowerCase();
    if (!searchStr) return { success: true, results: [] };

    var resultsMap = {};
    var sheetNames = ["등록 현황", "등록현황"]; // 원장님 요청: 등록현황 시트에서만 불러오기
    var sheetsSearched = [];
    var totalRowsProcessed = 0;

    sheetNames.forEach(function(sName) {
      var sheet = ss.getSheetByName(sName);
      if (!sheet) return;
      sheetsSearched.push(sName);

      var data = sheet.getDataRange().getValues();
      var cols = getRegColumnIndices(sheet); 
      totalRowsProcessed += data.length;

      for (var j = 1; j < data.length; j++) {
        var row = data[j];
        var foundInRow = false;

        // 전체 행 스캔 (성공했던 로직 유지: 어느 칸에 이름이 있든 검색)
        for (var k = 0; k < row.length; k++) {
          var cellVal = String(row[k] || "").trim().normalize("NFC");
          var cleanCell = cellVal.replace(/\s/g, "").toLowerCase();
          if (cleanCell.indexOf(searchStr) !== -1) {
            foundInRow = true;
            break;
          }
        }

        if (foundInRow) {
          var rName = String(row[cols.name] || row[1] || "").trim();
          var rPhoneRaw = String(row[cols.phone] || row[2] || "").trim();
          var rPhone = rPhoneRaw.replace(/[^0-9]/g, "");
          if (!rPhone) rPhone = "no-phone-" + sName + "-" + j;

          if (!resultsMap[rPhone]) {
            resultsMap[rPhone] = {
              name: rName || "이름없음",
              phone: rPhoneRaw || "-",
              activeList: [],
              closedList: []
            };
          }

          var rStatus = String(row[cols.status] || row[8] || "").trim();
          var expRaw = row[cols.expire] || row[6];
          var expStr = (expRaw instanceof Date) ? Utilities.formatDate(expRaw, "GMT+9", "yyyy-MM-dd") : String(expRaw || "-").split(" ")[0];

          var item = {
            membership: String(row[cols.membership] || row[4] || ""),
            remain: row[cols.remain] || row[7],
            expireDate: expStr,
            status: rStatus
          };

          if (rStatus.indexOf("진행") !== -1) resultsMap[rPhone].activeList.push(item);
          else resultsMap[rPhone].closedList.push(item);
        }
      }
    });

    var results = Object.values(resultsMap);
    if (results.length === 0) {
      return { 
        success: true, 
        results: [], 
        debug: "검색어: [" + searchStr + "], 시트: [" + sheetsSearched.join(", ") + "], 총 검사: " + totalRowsProcessed + "줄"
      };
    }

    // 마감 내역 정리 및 보너스 횟수 맵핑
    var bonusMap = {};
    var memberSheet = ss.getSheetByName("회원명단") || ss.getSheetByName("회원DB");
    if (memberSheet) {
      var mData = memberSheet.getDataRange().getValues();
      for (var i = 1; i < mData.length; i++) {
        var mPhone = String(mData[i][2] || "").replace(/[^0-9]/g, "");
        if (mPhone) bonusMap[mPhone] = mData[i][9]; 
      }
    }

    results.forEach(function(m) {
      var cleanPhone = m.phone.replace(/[^0-9]/g, "");
      m.bonusCount = bonusMap[cleanPhone] || 0;
      if (m.closedList.length > 1) m.closedList = [m.closedList[m.closedList.length - 1]];
    });

    return { success: true, results: results };
  } catch (e) {
    return { error: "검색 도중 치명적 오류: " + e.toString() };
  }
}

/**
 * [추천인 전용] 회원DB에서 성함으로 검색하여 추천인 후보를 리턴합니다.
 */
function searchAllMembers(nameStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName("회원명단");
    if (!memberSheet) return { error: "회원DB 시트가 없습니다." };
    
    var mData = memberSheet.getDataRange().getValues();
    var searchStr = String(nameStr || "").trim().toLowerCase();
    var results = [];
    
    if (searchStr.length < 1) return { success: true, results: [] };

    for (var i = 1; i < mData.length; i++) {
      var name = String(mData[i][1] || "").trim();
      var phone = String(mData[i][2] || "").trim();
      
      // 이름이 포함되어 있으면 결과에 추가
      if (name && name.toLowerCase().indexOf(searchStr) !== -1) {
        results.push({
          name: name,
          phone: phone
        });
      }
    }
    return { success: true, results: results };
  } catch (e) {
    return { error: "추천인 검색 오류: " + e.toString() };
  }
}

function manualAdminCheckIn(phoneStr, type) {
  // 관리자 수동 입실도 processAttendance와 동일한 로직을 사용하도록 연동합니다.
  // 다만 phoneStr이 전체 번호이므로 별도 처리가 필요 없습니다.
  return processAttendance(phoneStr, type, false);
}

// ──────────────────────────────────────────────
// 9. 자동 퇴실 트리거 및 퇴실 내역 수정 API
// ──────────────────────────────────────────────

function autoCheckoutJob() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    var resSheet = ss.getSheetByName("예약DB");
    if (!logSheet || !resSheet) return;

    var data = logSheet.getDataRange().getValues();
    var cols = getAttendanceColumnIndices(logSheet);
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    // 예약 데이터 미리 가져오기 (매칭용)
    var rData = resSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var logDateRaw = data[i][cols.date];
      var logDateStr = (logDateRaw instanceof Date) ? Utilities.formatDate(logDateRaw, "GMT+9", "yyyy-MM-dd") : String(logDateRaw).split(" ")[0];
      var status = String(data[i][cols.status]).trim();
      
      // 오늘 입실 상태인 로그만 검사
      if (logDateStr === todayStr && status === "입실") {
        var rowIdx = i + 1;
        var memberName = String(data[i][cols.name]);
        var reason = String(data[i][cols.reason]);
        var isCombo = reason.indexOf("복합") !== -1;
        var isJumping = reason.indexOf("점핑") !== -1 && !isCombo;
        var isTherapy = reason.indexOf("테라피") !== -1 && !isCombo;

        // 시간차 계산 (날짜(A) + 입실시간(B)을 합쳐서 정확한 입장 시각 계산)
        var inTimeStr = String(data[i][cols.inTime] || "");
        var entryDateTime = new Date(logDateRaw);
        
        if (inTimeStr.indexOf(":") !== -1) {
          var tParts = inTimeStr.split(":");
          entryDateTime.setHours(parseInt(tParts[0]), parseInt(tParts[1]), 0, 0);
        }

        var diffMs = now.getTime() - entryDateTime.getTime();
        var diffHours = diffMs / (1000 * 60 * 60);

        var shouldAutoCheckout = false;
        if (isJumping && diffHours >= 4) shouldAutoCheckout = true;
        if (isTherapy && diffHours >= 3) shouldAutoCheckout = true;
        if (isCombo && diffHours >= 5) shouldAutoCheckout = true;

        if (shouldAutoCheckout) {
          // 1. 예약DB에서 오늘 이 회원의 예약 시간 찾기 (지능형 매칭 - 이름+번호)
          var matchedClass = "";
          var memberPhoneStr = String(data[i][cols.phone]).replace(/[^0-9]/g, "");

          for (var k = 1; k < rData.length; k++) {
            var rDateRaw = rData[k][3];
            var rDateStr = (rDateRaw instanceof Date) ? Utilities.formatDate(rDateRaw, "GMT+9", "yyyy-MM-dd") : String(rDateRaw).split(" ")[0];
            var rPhone = String(rData[k][1]).replace(/[^0-9]/g, "");

            if (rDateStr === todayStr && String(rData[k][2]) === memberName && rPhone === memberPhoneStr) {
              matchedClass = String(rData[k][4]).substring(0, 5); // 입실시간(HH:mm)
              // 예약 상태가 '테라피중'이면 '완료'로 업데이트
              if (String(rData[k][9]).indexOf("테라피중") !== -1) {
                resSheet.getRange(k + 1, 10).setValue("완료");
                resSheet.getRange(k + 1, 12).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm:ss") + "[자동]");
              }
              break;
            }
          }

          // 2. 자동 퇴실 기록
          var autoTimeLog = (isJumping || isCombo) ? "1.0" : "";
          logSheet.getRange(rowIdx, cols.classes + 1).setValue(matchedClass); // J열: 참여클래스(예약시간)
          logSheet.getRange(rowIdx, cols.workoutTime + 1).setValue(autoTimeLog); // K열: 운동타임수
          logSheet.getRange(rowIdx, cols.status + 1).setValue("귀가[자동]"); // L열: 상태
          logSheet.getRange(rowIdx, cols.outTime + 1).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm") + "[자동]"); // M열: 퇴실시간
        }
      }
    }
  } catch (e) {
    console.error("자동 퇴실 처리 중 오류: " + e.toString());
  }
}

// 중복 함수 제거 (processAdminCheckout는 1057번 줄에 이미 정의되어 있음)

function editAdminCheckout(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    var memberSheet = ss.getSheetByName("회원명단");
    
    var rowIdx = data.rowIdx;
    var cols = getAttendanceColumnIndices(logSheet);
    var phoneStr = String(logSheet.getRange(rowIdx, cols.phone + 1).getValue());
    var currentExtraText = String(logSheet.getRange(rowIdx, cols.memo + 1).getValue());
    var wasExtraDeducted = currentExtraText.indexOf("(추가차감됨)") !== -1;
    
    // 로그 업데이트 (J:클래스, K:타임)
    logSheet.getRange(rowIdx, cols.classes + 1).setValue(data.classes);
    logSheet.getRange(rowIdx, cols.workoutTime + 1).setValue(data.timeLog);
    
    // 등록 현황 시트 준비
    var regSheet = ss.getSheetByName("등록 현황");
    var rData = regSheet.getDataRange().getValues();
    var rCols = getRegColumnIndices(regSheet);
    
    // 1. 등록 현황 시트에서 '진행중'인 회원권 찾기
    var targetRegIdx = -1;
    for (var i = 1; i < rData.length; i++) {
      var rPhone = String(rData[i][rCols.phone]).replace(/[^0-9]/g, "");
      var rStatus = String(rData[i][rCols.status] || "").trim();
      if (rPhone === phoneStr.replace(/[^0-9]/g, "") && rStatus.indexOf("진행") !== -1) {
        targetRegIdx = i + 1;
        break;
      }
    }
    
    if (targetRegIdx !== -1) {
      var remainCount = rData[targetRegIdx-1][rCols.remain];
      var isUnlimited = (remainCount === "(무제한)" || remainCount === "무제한");
      var curVal = isUnlimited ? 0 : Number(remainCount);
      var currentReason = String(logSheet.getRange(rowIdx, cols.reason + 1).getValue()).replace(" (추가차감)", "");

      if (!wasExtraDeducted && data.extraDeduct) {
        // 새로 체크됨: 1회 차감
        var newVal = isUnlimited ? "무제한" : (Number(curVal) - 1);
        if (!isUnlimited) regSheet.getRange(targetRegIdx, rCols.remain + 1).setValue(newVal);
        
        // 출석기록 업데이트 (G, H열 강제 주입)
        var currentDeduct = Number(logSheet.getRange("G" + rowIdx).getValue()) || -1;
        logSheet.getRange("G" + rowIdx).setValue(currentDeduct - 1); 
        logSheet.getRange("H" + rowIdx).setValue(newVal);
        
        logSheet.getRange(rowIdx, cols.reason + 1).setValue(currentReason + " (추가차감)");
        logSheet.getRange(rowIdx, cols.memo + 1).setValue("기준시간 초과 (1회 추가 차감됨)");
      } 
      else if (wasExtraDeducted && !data.extraDeduct) {
        // 체크 해제됨: 1회 환불
        var newVal = isUnlimited ? "무제한" : (Number(curVal) + 1);
        if (!isUnlimited) regSheet.getRange(targetRegIdx, rCols.remain + 1).setValue(newVal);
        
        // 출석기록 업데이트 (원복)
        var currentDeduct = Number(logSheet.getRange("G" + rowIdx).getValue()) || -2;
        logSheet.getRange("G" + rowIdx).setValue(currentDeduct + 1); 
        logSheet.getRange("H" + rowIdx).setValue(newVal);
        
        logSheet.getRange(rowIdx, cols.reason + 1).setValue(currentReason);
        logSheet.getRange(rowIdx, cols.memo + 1).setValue("");
      }
    } else if (data.extraDeduct) {
      // 회원권을 못 찾은 경우에도 메모는 남김 (경고 표시)
      logSheet.getRange(rowIdx, cols.memo + 1).setValue("⚠️회원권 못찾음 - 추가차감 실패 ★★★");
    }
    
    SpreadsheetApp.flush(); 
    
    var currentStatus = String(logSheet.getRange(rowIdx, cols.status + 1).getValue());
    if (currentStatus === "입실") {
      logSheet.getRange(rowIdx, 12).setValue("귀가");
      var now = new Date();
      var timeStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
      logSheet.getRange(rowIdx, 13).setValue(timeStr);
      
      // 예약DB 연동 (수정 시에도 귀가로 바뀌면 예약DB 업데이트)
      try {
        var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
        var mName = logSheet.getRange(rowIdx, 3).getValue();
        var mPhone = String(logSheet.getRange(rowIdx, 4).getValue()).replace(/[^0-9]/g, "");

        var resSheet = ss.getSheetByName("예약DB");
        var resData = resSheet.getDataRange().getDisplayValues();
        var todayParts = todayStr.match(/\d+/g);
        var todayNum = todayParts[0] + (todayParts[1].length === 1 ? "0" : "") + todayParts[1] + (todayParts[2].length === 1 ? "0" : "") + todayParts[2];

        for (var k = 1; k < resData.length; k++) {
          var rDateRaw = resData[k][3];
          var rDateParts = String(rDateRaw).match(/\d+/g);
          if (!rDateParts || rDateParts.length < 3) continue;
          var rDateNum = rDateParts[0] + (rDateParts[1].length === 1 ? "0" : "") + rDateParts[1] + (rDateParts[2].length === 1 ? "0" : "") + rDateParts[2];
          
          var rPhone = String(resData[k][1]).replace(/[^0-9]/g, "");
          var isPhoneMatch = (rPhone === mPhone) || (rPhone.length >= 8 && mPhone.length >= 8 && rPhone.slice(-8) === mPhone.slice(-8));

          if (rDateNum === todayNum && isPhoneMatch) {
            var status = String(resData[k][9]);
            if (status.indexOf("테라피") !== -1 || status.indexOf("예약") !== -1) {
              resSheet.getRange(k + 1, 10).setValue("귀가");
              resSheet.getRange(k + 1, 12).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm:ss"));
              break;
            }
          }
        }
      } catch(e) {}
    }

    return { success: true, message: "수정 내역이 안전하게 반영되었습니다." };
  } catch (e) { return { error: "수정 처리 오류: " + e.toString() }; }
}

function getClassMembersByDate(targetDateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    if (!logSheet) return { error: "시트 오류: '출석기록' 시트를 찾을 수 없습니다." };
    
    var data = logSheet.getDataRange().getValues();
    var classMembers = { "09시": [], "10시": [], "17시": [], "18시": [], "19시": [], "20시": [] };
    var allLogs = [];
    
    for (var i = 1; i < data.length; i++) {
      var timeRaw = data[i][1]; // B열: 입실시간
      var logDateStr = String(data[i][0]).split(" ")[0]; // A열: 날짜
      
      if (logDateStr === targetDateStr) {
        var name = String(data[i][2] || ""); // C열
        var phone = String(data[i][3] || ""); // D열
        var membership = String(data[i][4] || ""); // E열
        var classes = String(data[i][9] || ""); // J열
        var timeLog = String(data[i][10] || "1.0"); // K열
        var reason = String(data[i][8] || ""); // I열
        var status = String(data[i][11] || ""); // L열
        var extraDeduct = (reason.indexOf("추가차감") !== -1);

        allLogs.push({
          name: name,
          phone: phone,
          membership: membership,
          rowIdx: i + 1,
          classes: classes,
          timeLog: timeLog,
          extraDeduct: extraDeduct,
          status: status,
          inTime: String(timeRaw).indexOf(":") !== -1 ? String(timeRaw).split(":").slice(0, 2).join(":") : String(timeRaw),
          type: (reason.indexOf("테라피") !== -1 || reason.indexOf("보너스") !== -1) ? "테라피" : "점핑"
        });
        
        for (var slot in classMembers) {
          if (classes.indexOf(slot) !== -1) classMembers[slot].push(name);
        }
      }
    }
    return { success: true, classMembers: classMembers, allLogs: allLogs };
  } catch (e) { return { error: "명단 조회 오류: " + e.toString() }; }
}

// ──────────────────────────────────────────────
// 10. 신규 회원 가입 / 재등록 및 서명 API
// ──────────────────────────────────────────────

function getMemberRenewalData(phoneStr) {
  if (!phoneStr) return { error: "번호 없음" };
  var clean = String(phoneStr).replace(/[^0-9]/g, "");
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mData = ss.getSheetByName("회원명단").getDataRange().getValues();
    var rData = ss.getSheetByName("등록 현황").getDataRange().getValues();
    
    var res = { name: "-", phone: phoneStr, activeList: [] };
    
    // 1. 이름찾기
    for (var i=1; i<mData.length; i++) {
      if (String(mData[i][2]).replace(/[^0-9]/g, "") === clean) {
        res.name = String(mData[i][1]);
        break;
      }
    }
    
    // 2. 내역찾기
    for (var j=1; j<rData.length; j++) {
      if (String(rData[j][2]).replace(/[^0-9]/g, "") === clean) {
        var exp = rData[j][6];
        var expStr = (exp instanceof Date) ? (exp.getFullYear() + "-" + (exp.getMonth()+1) + "-" + exp.getDate()) : String(exp || "-");
        res.activeList.push({
          membership: String(rData[j][4] || ""),
          expireDate: expStr,
          remainCount: rData[j][7] || 0,
          status: String(rData[j][8] || "")
        });
      }
    }

    // 3. 설정
    var configs = [];
    var cData = ss.getSheetByName("설정").getDataRange().getValues();
    for (var k=1; k<cData.length; k++) {
      if (cData[k][0]) {
        configs.push({ 
          name: String(cData[k][0]), 
          count: Number(cData[k][1]) || 0,
          duration: Number(cData[k][2]) || 0, // 유효기간(일)
          price: Number(cData[k][6]) || 0 
        });
      }
    }
    
    return { success: true, member: res, config: configs };
  } catch (e) {
    return { error: e.toString() };
  }
}

function submitRegistration(data) {
  if (data.phone) data.phone = formatPhoneNumber(data.phone);
  if (data.referrerId) data.referrerId = formatPhoneNumber(data.referrerId);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName("회원명단");
    var regSheet = ss.getSheetByName("등록 현황");
    var configSheet = ss.getSheetByName("설정");
    
    var now = new Date();
    var logTimeStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm:ss");
    var startDateStr = data.startDate || Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    // 1. 설정 시트에서 권종 정보(가격, 횟수, 개월수) 가져오기
    var cData = configSheet.getDataRange().getValues();
    var passInfo = { baseCount: 0, validDays: 0, price: 0 };
    for (var c = 1; c < cData.length; c++) {
      if (cData[c][0] === data.membership) {
        passInfo.baseCount = Number(cData[c][1]) || 0;
        passInfo.validDays = Number(cData[c][2]) || 0; // 유효기간(일)
        passInfo.price = Number(cData[c][6]) || 0; // G열 가격
        break;
      }
    }
    
    // 2. 만료일 계산 로직 (일수 기준)
    var calcExpDate = function(baseDateStr, daysToAdd) {
      var d = new Date(baseDateStr);
      d.setDate(d.getDate() + Number(daysToAdd));
      // d.setDate(d.getDate() - 1); // 하루 빼기 로직 제거
      return Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
    };

    var mData = memberSheet.getDataRange().getValues();
    var existingRowIdx = -1;
    var existingExpDate = "";
    var existingRemain = 0;
    
    for (var i = 1; i < mData.length; i++) {
      if (String(mData[i][2]) === data.phone) {
        existingRowIdx = i + 1;
        existingExpDate = String(mData[i][7]); // H열 만료일
        existingRemain = mData[i][8]; // I열 잔여횟수
        break;
      }
    }
    
    var finalExpDate = "";
    var finalRemain = passInfo.baseCount;
    var finalMemo = ""; // 등록현황(장부)용 비고
    var payCategory = "신규등록"; // 기본값
    
    if (existingRowIdx !== -1) {
      // --- 재등록/연장 로직 ---
      var isNotExpired = false;
      if (existingExpDate && existingExpDate !== "undefined" && existingExpDate !== "-") {
        var expD = new Date(existingExpDate);
        expD.setHours(23, 59, 59, 999);
        if (now <= expD) isNotExpired = true;
      }
      
      payCategory = isNotExpired ? "연장결제" : "재결제"; 
      
      var gapDays = 0;
      if (existingExpDate && existingExpDate !== "undefined" && existingExpDate !== "-") {
        var prevExp = new Date(existingExpDate);
        gapDays = Math.floor((now.getTime() - prevExp.getTime()) / (1000 * 60 * 60 * 24));
      }

      var isMonthly = (data.membership.indexOf("월권") !== -1 || data.membership.indexOf("운동만") !== -1);
      
      // 연장 시: 기존 만료일이 남았으면 그 뒤로, 아니면 오늘부터 가산
      var baseDate = isNotExpired ? existingExpDate : startDateStr;
      finalExpDate = calcExpDate(baseDate, passInfo.validDays);
      
      if (isMonthly) {
        // 월권 연장 시: 합산하지 않고 새 횟수 부여 + 기존 횟수는 이월 딱지로
        finalRemain = passInfo.baseCount;
        if (curRemainVal > 0 && isNotExpired) {
          var carryTag = "[월권이월:" + curRemainVal + "회|기한:" + existingExpDate + "]";
          var currentMemo = String(memberSheet.getRange(existingRowIdx, 20).getValue() || "");
          var newDBMemo = currentMemo.replace(/\[월권이월:.*?\]/g, "").trim() + " " + carryTag;
          memberSheet.getRange(existingRowIdx, 20).setValue(newDBMemo.trim());
          finalMemo = "(" + startDateStr + ") " + curRemainVal + "회 이월딱지 생성 및 연장";
        } else {
          finalMemo = "(" + startDateStr + ") 만료 후 재결제";
        }
      } else {
        // 횟수권 연장 시: 기존 횟수와 합산
        finalRemain = curRemainVal + passInfo.baseCount;
        finalMemo = "(" + startDateStr + ") " + (isNotExpired ? "기존 잔여 " + curRemainVal + "회 합산 연장" : "만료 후 재결제");
      }
      
      // 회원DB 업데이트 (기존 회원 정보 갱신)
      memberSheet.getRange(existingRowIdx, 6).setValue(data.membership);
      memberSheet.getRange(existingRowIdx, 8).setValue(finalExpDate);
      memberSheet.getRange(existingRowIdx, 9).setValue(finalRemain);
      
      var oldHistory = String(memberSheet.getRange(existingRowIdx, 11).getValue() || "");
      var newHistoryEntry = startDateStr + ": " + data.membership + " (" + payCategory + ")";
      memberSheet.getRange(existingRowIdx, 11).setValue(oldHistory ? oldHistory + "\n" + newHistoryEntry : newHistoryEntry);

    } else {
      // --- 신규 가입 로직 ---
      finalExpDate = calcExpDate(startDateStr, passInfo.validDays);
      payCategory = "신규등록";
      finalMemo = "최초등록";
      
      var memberID = "M" + now.getTime();
      var firstHistory = startDateStr + ": " + data.membership + " 신규등록";
      
      var rowData = [
        memberID, data.name, data.phone, data.birthdate, data.address, 
        data.membership, startDateStr, finalExpDate, finalRemain, 0, 
        firstHistory, data.referrer, data.referrerId || "", data.goal, 
        data.goalWeight, data.medication, data.signature, "" // 20번 컬럼(비고/메모) 빈값
      ];
      memberSheet.appendRow(rowData);
    }
    
    // 3. 등록현황(매출장부) 정밀 기록 (12개 컬럼 매칭)
    // ["등록일", "이름", "휴대폰", "결제구분", "권종", "시작일", "만료일", "잔여횟수", "상태", "결제금액", "비고", "서명"]
    regSheet.appendRow([
      startDateStr,       // 1. 등록일
      data.name,           // 2. 이름
      data.phone,          // 3. 휴대폰
      payCategory,         // 4. 결제구분 (신규등록/재결제)
      data.membership,     // 5. 권종
      startDateStr,       // 6. 시작일
      finalExpDate,       // 7. 만료일
      passInfo.baseCount, // 8. 이번에 결제한 권종의 기본 횟수
      "진행중",            // 9. 상태
      passInfo.price,     // 10. 결제금액
      finalMemo,           // 11. 비고
      data.signature       // 12. 서명
    ]);
    
    // 4. 매출내역 자동 등록 (원장님 요청)
    submitSalesRecord({
      date: startDateStr,
      category: "회원등록",
      buyer: data.name,
      itemName: data.membership,
      amount: passInfo.price,
      payMethod: data.payMethod || "카드",
      memo: payCategory // "신규등록" 또는 "재결제"
    });

    // 5. 추천인 보너스 로직 (+1회 추가)
    if (data.referrerId) {
      for (var r = 1; r < mData.length; r++) {
        if (String(mData[r][2]).trim() === String(data.referrerId).trim()) {
          var refRow = r + 1;
          var currentBonus = Number(memberSheet.getRange(refRow, 10).getValue()) || 0;
          memberSheet.getRange(refRow, 10).setValue(currentBonus + 1);
          
          // 추천인 히스토리에도 기록
          var refOldHist = String(memberSheet.getRange(refRow, 11).getValue() || "");
          var refNewHist = startDateStr + ": [" + data.name + "]님 추천 보너스 +1회 적립";
          memberSheet.getRange(refRow, 11).setValue(refOldHist ? refOldHist + "\n" + refNewHist : refNewHist);
          break;
        }
      }
    }

    // 6. [원장님 요청] 문자발송 시트에 자동 기록 생성
    try {
      var smsSheet = ss.getSheetByName("문자발송");
      if (smsSheet) {
        var smsContent = generateSmsContent(data.name, payCategory, data.membership, finalExpDate, finalRemain, gapDays);
        smsSheet.appendRow([
          logTimeStr,    // 기록시간
          data.name,      // 이름
          data.phone,     // 전화번호
          payCategory,    // 안내분류 (신규등록/재결제)
          smsContent,     // 생성된문자내용
          "대기"          // 상태
        ]);
      }
    } catch(smsErr) {
      console.error("문자발송 기록 실패: " + smsErr.toString());
    }

    return { success: true, message: (existingRowIdx !== -1 ? "재등록 및 장부 기록이 완료되었습니다!" : "신규 가입 및 장부 기록이 완료되었습니다!") };
  } catch (e) {
    return { error: "처리 중 오류가 발생했습니다: " + e.toString() };
  }
}


/**
 * [공용] 알려주신 수식을 바탕으로 문자 메시지 내용을 생성하는 함수
 * @param {number} gapDays - 이전 만료일로부터 지난 일수 (재결제 시 구분용)
 */
function generateSmsContent(name, category, membership, expireDate, remainCount, gapDays) {
  var content = name + " 회원님, 노형점핑클럽입니다! ❤️\n\n";
  
  if (category === "신규등록") {
    content += "노형점핑과의 첫 만남을 진심으로 환영합니다! 😊\n" +
               "선택하신 [" + membership + "] 등록이 완료되었습니다.\n" +
               "📅 유효기간: ~ " + expireDate + "\n\n" +
               "회원님의 건강과 미모가 빛나도록 정성을 다해 돕겠습니다!";
  } else if (category === "연장결제") {
    content += "잊지 않고 연장 등록 해주셔서 감사합니다! ✨\n" +
               "기존 잔여분과 꼼꼼히 합산하여 [" + membership + "] 등록을 마쳤습니다.\n" +
               "📊 총 잔여횟수: " + remainCount + "회\n" +
               "📅 최종 만료일: ~ " + expireDate + "\n\n" +
               "꾸준한 관리가 최고의 결과를 만듭니다. 화이팅입니다! 🔥";
  } else if (category === "재결제") {
    if (gapDays > 30) {
      // 30일 이상 지난 경우 (오랜만에 오신 분)
      content += "다시 노형점핑을 찾아주셔서 정말 기뻐요! 🥰\n" +
                 "결제하신 [" + membership + "] 등록이 완료되었습니다.\n" +
                 "📅 유효기간: ~ " + expireDate + "\n\n" +
                 "다시 한번 믿고 등록해 주신 만큼, 정성을 다해 관리해 드릴게요! 💪";
    } else {
      // 30일 이내인 경우 (며칠 늦게 결제하신 분)
      content += "노형점핑을 다시 믿고 선택해 주셔서 감사합니다! 😊\n" +
                 "결제하신 [" + membership + "] 등록이 완료되었습니다.\n" +
                 "📅 유효기간: ~ " + expireDate + "\n\n" +
                 "이번에도 회원님의 건강한 변화를 위해 최선을 다하겠습니다! 🔥";
    }
  } else if (category === "추가결제") {
    content += "기존 프로그램과 더불어 [" + membership + "]을 추가해 주셔서 감사합니다! 🥰\n" +
               "회원님의 열정적인 도전을 보며 저희도 더 힘이 나네요.\n" +
               "📅 신규 권종 만료일: ~ " + expireDate + "\n\n" +
               "두 가지 프로그램 모두 시너지가 나도록 세심하게 관리해 드릴게요! ✨";
  } else {
    content += "결제하신 [" + membership + "] 등록이 완료되었습니다.\n" +
               "📅 유효기간: ~ " + expireDate + "\n\n" +
               "정성을 다해 관리해 드릴게요! 💪";
  }
  return content;
}




/**
 * [연장 전용] 회원 연장 / 재결제 처리
 */
function submitRenewal(data) {
  if (data.phone) data.phone = formatPhoneNumber(data.phone);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName("회원명단");
    var regSheet = ss.getSheetByName("등록 현황");
    var configSheet = ss.getSheetByName("설정");
    
    var now = new Date();
    var startDateStr = data.startDate;
    
    // 1. 설정 정보 가져오기
    var cData = configSheet.getDataRange().getValues();
    var passInfo = { baseCount: 0, validDays: 0, price: 0 };
    for (var c = 1; c < cData.length; c++) {
      if (cData[c][0] === data.membership) {
        passInfo.baseCount = Number(cData[c][1]) || 0;
        passInfo.validDays = Number(cData[c][2]) || 0; // 이제 '일' 단위로 처리
        passInfo.price = Number(cData[c][6]) || 0;
        break;
      }
    }

    // 2. 등록현황에서 '똑같은 이름'의 진행중 내역 찾기
    var regData = regSheet.getDataRange().getValues();
    var targetRowIdx = -1;
    var currentExp = startDateStr;
    var currentRemain = 0;
    var memberName = "";

    for (var k = regData.length - 1; k >= 1; k--) {
      if (String(regData[k][2]) === data.phone) {
        if (!memberName) memberName = regData[k][1];
        
        // 이전 방식: 이름이 정확히 똑같아야 했음
        // 수정 방식: 같은 "점핑 회수권" 계열이거나 "테라피 회수권" 계열이면 연장 대상으로 인정
        var isOngoing = (String(regData[k][8]).trim() === "진행중");
        var oldType = String(regData[k][4]).trim();
        var newType = data.membership.trim();
        
        var isSameFamily = (oldType === newType); // 1. 이름이 완전히 같은 경우
        
        // 2. 점핑 회수권 계열인지 확인 (예: 50회 -> 30회)
        if (!isSameFamily && oldType.indexOf("점핑") !== -1 && oldType.indexOf("회") !== -1 && newType.indexOf("점핑") !== -1 && newType.indexOf("회") !== -1) {
          isSameFamily = true;
        }
        
        // 3. 테라피 회수권 계열인지 확인
        if (!isSameFamily && oldType.indexOf("테라피") !== -1 && oldType.indexOf("회") !== -1 && newType.indexOf("테라피") !== -1 && newType.indexOf("회") !== -1) {
          isSameFamily = true;
        }

        if (isOngoing && isSameFamily) {
          targetRowIdx = k + 1;
          currentExp = String(regData[k][6]);
          currentRemain = Number(regData[k][7]) || 0;
          break;
        }
      }
    }

    // 4. 만료일 계산기 (일수 기준)
    var calcExpDate = function(baseDateStr, daysToAdd) {
      var d = new Date(baseDateStr);
      d.setDate(d.getDate() + Number(daysToAdd));
      // d.setDate(d.getDate() - 1); // 하루 빼기 로직 제거
      return Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
    };

    var finalExp = "";
    var finalRemain = passInfo.baseCount;
    var finalMemo = data.memo || "";
    
    // 이전 만료일로부터 지난 일수 계산
    var gapDays = 0;
    if (currentExp && currentExp !== "undefined" && currentExp !== "-") {
      var prevExp = new Date(currentExp);
      gapDays = Math.floor((now.getTime() - prevExp.getTime()) / (1000 * 60 * 60 * 24));
    }

    // 5. 연장 유형별 로직
    if (data.renewType === "연장결제") {
      // 연장결제: 기존 만료일이 남았으면 그 뒤로, 지났으면 오늘(시작일)부터
      var expD = new Date(currentExp);
      expD.setHours(23, 59, 59, 999);
      var baseDate = (now <= expD) ? currentExp : startDateStr;
      finalExp = calcExpDate(baseDate, passInfo.validDays);
      
      var isMonthly = (data.membership.indexOf("월권") !== -1 || data.membership.indexOf("운동만") !== -1);
      
      if (isMonthly) {
        // 월권 연장 시: 숫자를 합산하지 않고 '이월 딱지' 시스템 사용
        finalRemain = passInfo.baseCount;
        var isNotExpired = (now <= expD);
        
        if (currentRemain > 0 && isNotExpired) {
          var carryTag = "[월권이월:" + currentRemain + "회|기한:" + currentExp + "]";
          finalMemo = "잔여 " + currentRemain + "회 이월딱지 생성 (" + currentExp + "까지)";
          
          // 회원DB의 메모(T열)에 이월 딱지 삽입
          var mDataForTag = memberSheet.getDataRange().getValues();
          for (var i = 1; i < mDataForTag.length; i++) {
            if (String(mDataForTag[i][2]) === data.phone) {
              var mRowTag = i + 1;
              var oldTMemo = String(memberSheet.getRange(mRowTag, 20).getValue() || "");
              var newTMemo = oldTMemo.replace(/\[월권이월:.*?\]/g, "").trim() + " " + carryTag;
              memberSheet.getRange(mRowTag, 20).setValue(newTMemo.trim());
              break;
            }
          }
        } else {
          finalMemo = "만료 후 연장결제 (이월 없음)";
        }
      } else {
        // 횟수권 연장 시: 기존 횟수와 합산
        finalRemain = currentRemain + passInfo.baseCount;
        finalMemo = "잔여 " + currentRemain + "회 합산 연장 (" + finalRemain + "회)";
      }
      
      // 기존 내역 처리 (원장님 요청: 기존건 0으로 만들고 상태 변경)
      if (targetRowIdx !== -1) {
        var oldMemo = String(regSheet.getRange(targetRowIdx, 11).getValue() || "");
        regSheet.getRange(targetRowIdx, 9).setValue("만료(연장)"); 
        regSheet.getRange(targetRowIdx, 8).setValue(0);
        regSheet.getRange(targetRowIdx, 11).setValue(oldMemo + " / 잔여횟수 " + currentRemain + "회 연장결제로 인해 차기이월로 0회 처리");
      }
    } else if (data.renewType === "재결제") {
      // 재결제: 오늘(시작일)부터 설정된 기간/횟수만 적용
      finalExp = calcExpDate(startDateStr, passInfo.validDays);
      finalRemain = passInfo.baseCount;
      finalMemo = "기존 " + currentRemain + "회 소멸 후 재결제";
      // 기존 내역은 그냥 마감 처리
      if (targetRowIdx !== -1) {
        regSheet.getRange(targetRowIdx, 9).setValue("마감(재결제)");
      }
    } else if (data.renewType === "추가결제") {
      // 추가결제: 기존 내역과 별개로 오늘(시작일)부터 설정된 기간/횟수 적용
      finalExp = calcExpDate(startDateStr, passInfo.validDays);
      finalRemain = passInfo.baseCount;
      finalMemo = "추가 구매 (별도 기간/횟수 적용)";
    } else {
      // 기타 (혹시 모를 예외 처리)
      finalExp = calcExpDate(startDateStr, passInfo.validDays);
      finalRemain = passInfo.baseCount;
      finalMemo = "기타 결제";
    }

    // 6. 회원DB 업데이트 (요약 정보 업데이트)
    var mData = memberSheet.getDataRange().getValues();
    for (var i = 1; i < mData.length; i++) {
      if (String(mData[i][2]) === data.phone) {
        var mRow = i + 1;
        memberSheet.getRange(mRow, 6).setValue(data.membership);
        memberSheet.getRange(mRow, 8).setValue(finalExp);
        memberSheet.getRange(mRow, 9).setValue(finalRemain);
        var oldHist = String(memberSheet.getRange(mRow, 11).getValue() || "");
        memberSheet.getRange(mRow, 11).setValue(oldHist + "\n" + startDateStr + ": " + data.membership + " (" + data.renewType + ")");
        break;
      }
    }
    
    // 7. 등록현황(장부) 새 행 추가
    regSheet.appendRow([
      startDateStr,       // 1. 등록일
      memberName,         // 2. 이름
      data.phone,          // 3. 휴대폰
      data.renewType,      // 4. 결제구분
      data.membership,     // 5. 권종
      startDateStr,       // 6. 시작일
      finalExp,           // 7. 만료일
      finalRemain,        // 8. 잔여횟수(합산 결과)
      "진행중",            // 9. 상태
      passInfo.price,     // 10. 결제금액 (횟수가 아닌 시트의 실제 가격을 기록)
      finalMemo,           // 11. 비고
      data.signature       // 12. 서명
    ]);

    // 8. 매출내역 자동 등록 (원장님 요청)
    submitSalesRecord({
      date: startDateStr,
      category: "회원등록",
      buyer: memberName,
      itemName: data.membership,
      amount: data.price,
      payMethod: data.payMethod || "카드",
      memo: (data.renewType === "연장결제" || data.renewType === "재결제") ? "재등록" : "추가등록"
    });

    // 9. [원장님 요청] 문자발송 시트에 자동 기록 생성
    try {
      var smsSheet = ss.getSheetByName("문자발송");
      if (smsSheet) {
        var smsContent = generateSmsContent(memberName, data.renewType, data.membership, finalExp, finalRemain, gapDays);
        smsSheet.appendRow([
          Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm:ss"),
          memberName,
          data.phone,
          data.renewType,
          smsContent,
          "대기"
        ]);
      }
    } catch(smsErr) {
      console.error("문자발송 기록 실패: " + smsErr.toString());
    }



    return { success: true, message: "연장 결제 처리가 범주별로 정확히 합산되었습니다!" };

  } catch (e) {
    return { error: "연장 처리 오류: " + e.toString() };
  }
}


function submitWorkLog(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var workLogSheet = ss.getSheetByName("업무일지");
    var dateStr = data.date || Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    
    var wlData = workLogSheet.getDataRange().getValues();
    var targetRowIdx = -1;
    
    // 날짜가 같은 행이 있는지 확인 (업데이트용)
    for (var i = 1; i < wlData.length; i++) {
      var rowDateRaw = wlData[i][0];
      var rowDateStr = (rowDateRaw instanceof Date) ? Utilities.formatDate(rowDateRaw, "GMT+9", "yyyy-MM-dd") : String(rowDateRaw).split(" ")[0];
      if (rowDateStr === dateStr) {
        targetRowIdx = i + 1;
        break;
      }
    }

    var rowValues = [
      dateStr,
      data.author,
      data.jumpingList,
      data.muscleList,
      data.stats["09시"],
      data.stats["10시"],
      data.stats["17시"],
      data.stats["18시"],
      data.stats["19시"],
      data.stats["20시"],
      data.stats.total,
      data.stats.jumping,
      data.stats.therapy,
      data.remarks,
      data.issues
    ];

    if (targetRowIdx !== -1) {
      // 기존 행 업데이트
      workLogSheet.getRange(targetRowIdx, 1, 1, rowValues.length).setValues([rowValues]);
      return { success: true, message: dateStr + " 업무일지가 업데이트되었습니다." };
    } else {
      // 새 행 추가
      workLogSheet.appendRow(rowValues);
      return { success: true, message: dateStr + " 업무일지가 새로 등록되었습니다." };
    }
  } catch (e) {
    return { error: e.toString() };
  }
}

function getWorkLogHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var workLogSheet = ss.getSheetByName("업무일지");
    if (!workLogSheet) return { error: "업무일지 시트가 없습니다." };
    
    var data = workLogSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, logs: [] };
    
    var logs = [];
    // 최근 15개만 가져오기 (역순)
    var start = Math.max(1, data.length - 15);
    for (var i = data.length - 1; i >= start; i--) {
      var row = data[i];
      
      // ★ 날짜가 Date 객체이면 문자열로 변환 (이걸 안 하면 전송 시 멈춤)
      var dateVal = row[0];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, "GMT+9", "yyyy-MM-dd");
      } else {
        dateVal = String(dateVal);
      }
      
      logs.push({
        date: dateVal,
        author: String(row[1] || ""),
        jumpingList: String(row[2] || ""),
        muscleList: String(row[3] || ""),
        stats: {
          "09시": String(row[4] || "0"), "10시": String(row[5] || "0"), 
          "17시": String(row[6] || "0"), "18시": String(row[7] || "0"), 
          "19시": String(row[8] || "0"), "20시": String(row[9] || "0"),
          total: String(row[10] || "0"),
          jumping: String(row[11] || "0"),
          therapy: String(row[12] || "0")
        },
        remarks: String(row[13] || ""),
        issues: String(row[14] || "")
      });
    }
    return { success: true, logs: logs };
  } catch (e) {
    return { error: "업무일지 조회 오류: " + e.toString() };
  }
}

// ──────────────────────────────────────────────
// 12. 판매내역 등록/조회 API
// ──────────────────────────────────────────────

function submitSalesRecord(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName("판매내역");
    if (!salesSheet) return { error: "판매내역 시트가 없습니다." };
    
    var now = new Date();
    var dateStr = data.date || Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, "GMT+9", "HH:mm:ss");
    var fullDateStr = dateStr + " " + timeStr;
    var idStr = "S" + dateStr.replace(/-/g, "") + Utilities.formatDate(now, "GMT+9", "HHmmss");
    
    salesSheet.appendRow([
      idStr,
      fullDateStr,
      data.category,
      data.buyer,
      data.itemName,
      Number(data.amount) || 0,
      data.payMethod,
      data.memo || ""
    ]);
    
    return { success: true, message: dateStr + " 매출이 등록되었습니다. (" + data.category + " / " + Number(data.amount).toLocaleString() + "원)" };
  } catch (e) {
    return { error: "매출 등록 오류: " + e.toString() };
  }
}

/**
 * 설정 시트에서 회원권 종류와 가격 정보를 가져옵니다.
 */
function getMembershipConfig() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var configSheet = ss.getSheetByName("설정");
    var data = configSheet.getDataRange().getValues();
    var config = [];
    for (var i = 1; i < data.length; i++) {
      config.push({
        name: data[i][0],
        price: data[i][6] || 0
      });
    }
    return { success: true, config: config };
  } catch (e) { return { error: e.toString() }; }
}

function getSalesByDate(dateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName("판매내역");
    if (!salesSheet) return { error: "판매내역 시트가 없습니다." };
    
    var data = salesSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, sales: [], totalAmount: 0, summary: {} };
    
    var sales = [];
    var totalAmount = 0;
    var summary = {}; // 구분별 합계
    
    for (var i = 1; i < data.length; i++) {
      var rowDateRaw = data[i][1]; // B열: 날짜
      var rowDateStr = "";
      if (rowDateRaw instanceof Date) {
        rowDateStr = Utilities.formatDate(rowDateRaw, "GMT+9", "yyyy-MM-dd");
      } else {
        rowDateStr = String(rowDateRaw).split(" ")[0].split("T")[0];
      }
      
      if (rowDateStr === dateStr) {
        var category = String(data[i][2] || "");
        var amount = Number(data[i][5]) || 0;
        
        sales.push({
          id: String(data[i][0] || ""),
          date: rowDateStr,
          category: category,
          buyer: String(data[i][3] || ""),
          itemName: String(data[i][4] || ""),
          amount: amount,
          payMethod: String(data[i][6] || ""),
          memo: String(data[i][7] || "")
        });
        
        totalAmount += amount;
        summary[category] = (summary[category] || 0) + amount;
      }
    }
    
    // summary를 배열로 변환 (직렬화 안전)
    var summaryArr = [];
    for (var key in summary) {
      summaryArr.push({ category: key, amount: summary[key] });
    }
    
    return { success: true, sales: sales, totalAmount: totalAmount, summary: summaryArr };
  } catch (e) {
    return { error: "매출 조회 오류: " + e.toString() };
  }
}

function getTodaySales() {
  var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
  return getSalesByDate(todayStr);
}

// ──────────────────────────────────────────────
// 13. 관리자 전용 상세 출석 관리 API
// ──────────────────────────────────────────────

/**
 * 특정 날짜의 출석 인원 명단을 상세 정보(rowIdx 포함)와 함께 가져옵니다.
 */

function getTodayWorkLogStats(dateVal) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    if (!logSheet) return { error: "출석기록 시트가 없습니다." };

    var data = logSheet.getDataRange().getValues();
    var cols = getAttendanceColumnIndices(logSheet);
    
    var stats = { "09시": 0, "10시": 0, "17시": 0, "18시": 0, "19시": 0, "20시": 0, "total": 0, "jumping": 0, "therapy": 0 };

    for (var i = 1; i < data.length; i++) {
      var dateRaw = data[i][cols.date];
      if (!dateRaw) continue;
      var logDateStr = (dateRaw instanceof Date) ? Utilities.formatDate(dateRaw, "GMT+9", "yyyy-MM-dd") : String(dateRaw).split(" ")[0];
      
      if (logDateStr === dateVal) {
        stats.total++; // 중복 제거 없이 모든 출석행 카운트

        var classes = String(data[i][cols.classes] || "");
        var type = String(data[i][cols.type] || "");
        var reason = String(data[i][cols.reason] || "");
        
        // 시간대별 집계
        for (var slot in stats) {
          if (slot.indexOf("시") !== -1 && classes.indexOf(slot) !== -1) {
            stats[slot]++;
          }
        }

        // 점핑/테라피/복합 판단
        var isCombo = (type === "복합" || reason.indexOf("복합") !== -1);
        var isTherapy = (type.indexOf("테라피") !== -1 || reason.indexOf("테라피") !== -1 || reason.indexOf("보너스") !== -1) && !isCombo;
        var isJumping = !isTherapy && !isCombo;

        if (isCombo) {
          stats.jumping++;
          stats.therapy++;
        } else if (isTherapy) {
          stats.therapy++;
        } else {
          stats.jumping++;
        }
      }
    }
    // 업무일지 시트에서 기존 기록이 있는지 확인
    var existingLog = null;
    var workLogSheet = ss.getSheetByName("업무일지");
    if (workLogSheet) {
      var wlData = workLogSheet.getDataRange().getValues();
      for (var j = 1; j < wlData.length; j++) {
        var rowDateRaw = wlData[j][0];
        var rowDateStr = (rowDateRaw instanceof Date) ? Utilities.formatDate(rowDateRaw, "GMT+9", "yyyy-MM-dd") : String(rowDateRaw).split(" ")[0];
        if (rowDateStr === dateVal) {
          existingLog = {
            author: wlData[j][1],
            jumpingList: wlData[j][2],
            muscleList: wlData[j][3],
            remarks: wlData[j][13],
            issues: wlData[j][14]
          };
          break;
        }
      }
    }

    return { success: true, stats: stats, existingLog: existingLog };
  } catch (e) {
    return { error: "통계 계산 오류: " + e.toString() };
  }
}
/**
 * 특정 날짜의 출석 명단을 타임별/전체로 분류하여 상세 정보와 함께 가져옵니다.
 */
function getClassMembersByDate(dateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    if (!logSheet) return { error: "출석기록 시트가 없습니다." };

    var data = logSheet.getDataRange().getDisplayValues(); // 시간 형식을 위해 DisplayValues 사용
    var cols = getAttendanceColumnIndices(logSheet);
    
    var classMembers = { "09시": [], "10시": [], "17시": [], "18시": [], "19시": [], "20시": [] };
    var allLogs = [];

    for (var i = 1; i < data.length; i++) {
      var logDateStr = data[i][cols.date].split(" ")[0]; // 이미 문자열

      if (logDateStr === dateStr) {
        var mName = String(data[i][cols.name]);
        var mInTime = String(data[i][cols.inTime] || data[i][1] || ""); // 입실시간 (B열 또는 인덱스)
        var mClasses = String(data[i][cols.classes] || "");
        var mType = String(data[i][cols.type] || "");
        var mReason = String(data[i][cols.reason] || "");
        var mTimeLog = String(data[i][cols.workoutTime] || "");
        var mStatus = String(data[i][cols.status] || "");
        
        // 타임별 명단 분류
        for (var slot in classMembers) {
          if (mClasses.indexOf(slot) !== -1) {
            classMembers[slot].push(mName);
          }
        }

        // 전체 로그 객체 생성
        allLogs.push({
          name: mName,
          inTime: mInTime,   // 입실 시간 추가
          membership: mType, // E열
          type: mType,       // 시각적 구분을 위한 유형
          classes: mClasses, // J열 (참여타임)
          timeLog: mTimeLog, // K열 (운동타임수)
          change: String(data[i][cols.change] || "0").replace("-", ""), // G열 (차감횟수)
          status: mStatus,
          reason: mReason
        });
      }
    }

    return { success: true, classMembers: classMembers, allLogs: allLogs };
  } catch (e) {
    return { error: "명단 조회 오류: " + e.toString() };
  }
}

/**
 * 33 챌린지 인바디 기록 저장 (기존 '인바디 입력' 시트와 호환)
 */
function saveInBodyRecord(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("인바디 입력");
    if (!sheet) return { error: "'인바디 입력' 시트가 없습니다. 시트명을 확인해주세요." };
    
    // 기존 구조에 맞춰 빈 행 추가 (A열 비우기, B 측정일, C 성함, D ID, E 체중, F 근량, G 지방, H 메모)
    sheet.appendRow([
      "", // A
      data.date || Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd"), // B (index 1)
      data.name, // C (index 2)
      data.phone, // D (index 3)
      data.weight, // E (index 4)
      data.muscle, // F (index 5)
      data.fat, // G (index 6)
      data.memo || "" // H (index 7)
    ]);
    
    return { success: true, message: data.name + " 님의 인바디 기록이 저장되었습니다." };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * 33 챌린지 전체 기록 조회 ('인바디 입력' 시트 기준)
 */
function getInBodyHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("인바디 입력");
    if (!sheet) return { success: true, records: [] };
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, records: [] };
    
    var records = [];
    // 최근 20개 정도만 가져오기
    var start = Math.max(1, data.length - 20);
    for (var i = data.length - 1; i >= start; i--) {
      var row = data[i];
      var dateVal = row[1]; // B열
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, "GMT+9", "yyyy-MM-dd");
      }
      
      records.push({
        date: String(dateVal || ""),
        name: String(row[2] || ""), // C열
        phone: String(row[3] || ""), // D열
        weight: row[4], // E
        muscle: row[5], // F
        fat: row[6], // G
        memo: String(row[7] || ""), // H
        rowIdx: i + 1
      });
    }
    return { success: true, records: records };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * 33 챌린지 성적 및 랭킹 조회 (기존 '33챌린지 랭킹뷰' 시트 기준)
 */
function getChallengeRanking() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지 랭킹뷰");
    if (!sheet) return { error: "'33챌린지 랭킹뷰' 시트가 없습니다. 노트북에서 정산 함수를 먼저 실행해주세요." };
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, rankings: [] };
    
    var rankings = [];
    // 원장님께서 주신 랭킹뷰 순서:
    // A: 정산기준(0), B: 회원ID(1), C: 이름(2), D: 시즌명(3) ... K: 주간체지방변화율(10), L: 시즌체지방변화율(11), M: 상담메모(12)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[2]) continue; // 이름이 없으면 패스
      
      rankings.push({
        weekLabel: String(row[0] || ""), // 정산기준 (주차 정보)
        name: String(row[2] || ""), // 이름
        score: row[11] || 0 // 시즌 체지방 변화율 (성적 기준)
      });
    }
    
    // 성적(변화율) 순으로 정렬 (체지방은 많이 빠질수록 성적이 좋으므로 오름차순 또는 내림차순 선택 필요)
    // 보통 변화율이 - 값일 것이므로 숫자가 작을수록(마이너스가 클수록) 상위권
    rankings.sort((a, b) => {
      var scoreA = parseFloat(a.score) || 0;
      var scoreB = parseFloat(b.score) || 0;
      return scoreA - scoreB; 
    });

    return { success: true, rankings: rankings };
  } catch (e) { return { error: e.toString() }; }
}

/**
 * 상단 메뉴 생성 (기존 대시보드 도구 통합)
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 노형점핑 대시보드')
      .addItem('💌 [문자발송(토)] 잔여현황 안내 팝업', 'showMemberReportPopup')
      .addItem('📧 [마우스드래그] 건강 리포트 발송 팝업', 'showIntegratedReportPopup') 
      .addSeparator()
      .addItem('✔️ 주간 잔여횟수 현황 누적(매주 일 5시 트리거)', 'backupRemainingSessions')
      .addSeparator()
      .addItem('📊1. 주간 통계 생성(회원명단 병기)', 'runWeeklyStats')
      .addItem('📊2. 월간 통계 생성(회원명단 병기)', 'runMonthlyStats')
      .addSeparator()
      .addItem('❇️33 챌린지 주별 누적 정산 시스템', 'run33ChallengeWeeklySettlement') 
      .addSeparator()
      .addItem('🎉각종 순위표 카드이미지 만들고 저장하기', 'showRankingCardBySelection') 
      .addItem('👥 챌린지 카드 (단톡방용-실명)', 'showChallengeCard_Full')
      .addItem('📱 챌린지 카드 (인스타용-숨김)', 'showChallengeCard_Masked')
      .addSeparator()
      .addItem('⚙️ ERP 시트 구조 최적화 (강제 업데이트)', 'forceUpdateAllHeaders')
      .addToUi();
}

/**
 * 노형점핑클럽 - 전체 안내 및 잔여 현황 발송 팝업
 */
function showMemberReportPopup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("등록 현황");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("⚠️ [등록 현황] 시트가 없습니다.");
    return;
  }

  var data = sheet.getDataRange().getValues();
  var list = ""; var allPhones = []; var rowIndices = []; var count = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][10] === "진행중"){
      var name = data[i][3];      // D: 이름
      var phone = String(data[i][4]).replace(/[^0-9]/g, ""); // E: 휴대폰
      var product = data[i][6];   // G: 권종
      var expiry = Utilities.formatDate(new Date(data[i][8]), "GMT+9", "yyyy-MM-dd"); // I: 만료일
      var remains = data[i][9];   // J: 잔여횟수

      var msg = name + " 회원님, 노형점핑클럽입니다! ❤️\n\n" +
                " 현재 잔여 현황을 안내드립니다. 😊\n\n" +
                "🏃 이용권: [" + product + "]\n" +
                "📊 남은횟수: " + remains + "회\n" +
                "📅 만료일자: ~ " + expiry + "\n\n" +
                "잔여 현황이 맞지 않으시면 연락주세요.^^\n" +"더 건강해지고 예뻐지시도록 노형점핑이 함께 할게요!\n" +" 즐거운 주말 보내시고 다음주에 클럽에서 뵙겠습니다. ✨";

      allPhones.push(phone);
      rowIndices.push(i + 1);

      var safeMsg = msg.replace(/\\/g, "\\\\").replace(/'/g, "\\'").replace(/"/g, '\\"').replace(/\n/g, "\\n");

      list += '<div id="row_' + (i+1) + '" style="margin-bottom:12px; border:1px solid #dee2e6; padding:15px; border-radius:10px; background:white; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">' +
              '<div style="font-size:15px; margin-bottom:8px;"><b>👤 ' + name + ' 회님</b> <span style="color:#868e96; font-size:12px;">(' + phone + ')</span></div>' +
              '<div style="font-size:13px; color:#495057; background:#f8f9fa; padding:10px; border-radius:5px; margin-bottom:10px; line-height:1.5; white-space:pre-wrap;">' + msg + '</div>' +
              '<div style="display:flex; gap:6px;">' +
              '<button onclick="copyT(\'' + phone + '\', this)" style="flex:1; padding:10px; border:1px solid #ced4da; border-radius:6px; background:white; cursor:pointer;">📞 번호 복사</button>' +
              '<button onclick="copyT(\'' + safeMsg + '\', this)" style="flex:2; padding:10px; border:none; border-radius:6px; background:#e3fafd; color:#0b7285; font-weight:bold; cursor:pointer;">💬 문구 복사</button>' +
              '<button onclick="markDone(' + (i+1) + ', this)" style="flex:1; padding:10px; border:none; border-radius:6px; background:#f1f3f5; cursor:pointer;">✅ 완료</button>' +
              '</div></div>';
      count++;
    }
  }

  var script = '<script>' +
    'function copyT(t,b){var a=document.body.appendChild(document.createElement("textarea"));a.value=t;a.select();document.execCommand("copy");document.body.removeChild(a);b.innerText="복사됨!";b.style.background="#fab005";b.style.color="white";setTimeout(function(){b.innerText= (t.length<15?"📞 번호":"💬 문구")+" 복사"; b.style.background=(t.length<15?"white":"#e3fafd"); b.style.color=(t.length<15?"black":"#0b7285");},800);}' +
    'function markDone(r,b){google.script.run.withSuccessHandler(function(){document.getElementById("row_"+r).style.background="#ebfbee"; document.getElementById("row_"+r).style.opacity="0.6"; b.innerText="확인됨"; b.disabled=true;}).setCheck(r);}' +
    'function markAllDone(rs){ if(confirm("선택한 모든 회원을 완료 처리할까요?")){ google.script.run.withSuccessHandler(function(){google.script.host.close();}).setAllChecks(rs); } }</script>';

  var header = count > 0 ?
    '<div style="background:#fff4e6; padding:15px; border:2px solid #fd7e14; border-radius:10px; margin-bottom:20px; text-align:center;">' +
    '<b>📢 회원 잔여현황 안내 (' + count + '건)</b><br>' +
    '<textarea id="allP" style="width:100%; margin:10px 0; padding:5px; height:50px;">' + allPhones.join(", ") + '</textarea>' +
    '<button onclick="copyT(document.getElementById(\'allP\').value, this)" style="width:100%; padding:10px; background:#fd7e14; color:white; border:none; border-radius:6px; font-weight:bold; cursor:pointer;">전체 번호 복사</button>' +
    '<button onclick="markAllDone([' + rowIndices.join(",") + '])" style="width:100%; padding:10px; background:#40c057; color:white; border:none; border-radius:6px; margin-top:5px; font-weight:bold; cursor:pointer;">전체 완료 체크</button></div>'
    : "<div style='padding:50px; text-align:center;'>대상자가 없습니다. 😊</div>";

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<div style="font-family:sans-serif; background:#f8f9fa; padding:15px; min-height:100%;">' + script + header + list + '</div>').setWidth(450).setHeight(700), "🧡 노형점핑 회원 관리 리포트");
}

function setCheck(r) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("등록 현황");
  sheet.getRange(r, 14).setValue(true); 
}

function setAllChecks(rs) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("등록 현황");
  rs.forEach(function(r) { sheet.getRange(r, 14).setValue(true); });
}

function showIntegratedReportPopup() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var response = ui.prompt("📊 (1/3) 단계", "데이터가 담긴 시트명을 입력해 주세요.", ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() != ui.Button.OK) return;
  var targetSheetName = response.getResponseText().trim();
  var sheet = ss.getSheetByName(targetSheetName); 
  if (!sheet) { ui.alert("⚠️ 시트를 찾을 수 없습니다."); return; }

  var periodResponse = ui.prompt("🗓️ (2/3) 단계", "문자에 표시할 리포트 기간을 입력해 주세요.\n(예: 2월 3주차)", ui.ButtonSet.OK_CANCEL);
  if (periodResponse.getSelectedButton() != ui.Button.OK) return;
  var periodName = periodResponse.getResponseText().trim();

  var dateResponse = ui.prompt("📅 (3/3) 단계", "인바디 매칭 기준 날짜를 입력하세요.\n(예: 2/21)", ui.ButtonSet.OK_CANCEL);
  if (dateResponse.getSelectedButton() != ui.Button.OK) return;
  var inputDateStr = dateResponse.getResponseText().trim();

  var currentTargetLabel = getTargetSeasonLabelByInput(inputDateStr);
  if (!currentTargetLabel || currentTargetLabel === "") {
    ui.alert("⚠️ 해당 날짜에 맞는 시즌 정보를 찾지 못했습니다.");
    return;
  }

  var weeklyListSheet = ss.getSheetByName("33챌린지 주별 누적");
  var weeklyData = weeklyListSheet ? weeklyListSheet.getDataRange().getValues() : [];
  var range = ss.getActiveRange();
  var data = range.getValues();
  var startRow = range.getRow();
  var list = ""; var count = 0;

  for (var i = 0; i < data.length; i++) {
    var actualRow = startRow + i;
    var id = String(data[i][0] || "").trim();      
    var name = data[i][1];                         
    var phone = String(data[i][2]).replace(/[^0-9]/g, ""); 
    var goal = data[i][9] || "건강관리";            
    var weekCount = data[i][16] || 0;               
    var weekTims = data[i][17] || 0;                
    var weekTimeStr = data[i][18] || "0시간 0분";    
    var monthCount = data[i][19] || 0;              
    var monthTims = data[i][20] || 0;               
    var monthTimeStr = data[i][21] || "0시간 0분";  
    var weight = data[i][22] || "-";                
    var fat = data[i][23] || "-";                   
    var muscle = data[i][24] || "-";                
    var score = data[i][25] || "-";                 

    // 매칭 인바디 찾기
    var inBodyNote = "인바디 매칭 정보 없음";
    for(var w=0; w<weeklyData.length; w++){
      if(String(weeklyData[w][0]) === id){
        inBodyNote = "인바디 변화 기록 있음"; break;
      }
    }

    var msg = name + " 회원님, 33챌린지 " + periodName + " 리포트입니다! 🏆\n\n" +
              "🔥 주간 활동 요약\n" +
              "🏃 출석: " + weekCount + "회 / 운동: " + weekTimeStr + "\n\n" +
              "📅 한달 누적 현황\n" +
              "✨ 출석: " + monthCount + "회 / 운동: " + monthTimeStr + "\n\n" +
              "📊 신체 변화 (매칭기준: " + inputDateStr + ")\n" +
              "⚖️ 체중: " + weight + "kg / 체지방: " + fat + "%\n" +
              "💪 골격근: " + muscle + "kg / 점수: " + score + "점\n\n" +
              "목표하신 [" + goal + "]를 향해 잘 가고 계십니다. 이번주도 화이팅하세요! ❤️";

    var safeMsg = msg.replace(/\\/g, "\\\\").replace(/'/g, "\\'").replace(/"/g, '\\"').replace(/\n/g, "\\n");

    list += '<div id="row_' + actualRow + '" style="margin-bottom:12px; border:1px solid #dee2e6; padding:15px; border-radius:10px; background:white;">' +
            '<div style="font-size:15px; margin-bottom:8px;"><b>👤 ' + name + ' 회님</b> <span style="color:#868e96; font-size:12px;">(' + phone + ')</span></div>' +
            '<div style="font-size:13px; color:#495057; background:#f8f9fa; padding:10px; border-radius:5px; margin-bottom:10px; white-space:pre-wrap;">' + msg + '</div>' +
            '<div style="display:flex; gap:6px;">' +
            '<button onclick="copyT(\'' + phone + '\', this)" style="flex:1; padding:10px;">📞 번호</button>' +
            '<button onclick="copyT(\'' + safeMsg + '\', this)" style="flex:2; padding:10px;">💬 문구 복사</button>' +
            '<button onclick="markDone(' + actualRow + ', this)" style="flex:1; padding:10px;">✅ 완료</button>' +
            '</div></div>';
    count++;
  }

  var script = '<script>' +
    'function copyT(t,b){var a=document.body.appendChild(document.createElement("textarea"));a.value=t;a.select();document.execCommand("copy");document.body.removeChild(a);b.innerText="복사됨!";setTimeout(function(){b.innerText= (t.length<15?"📞 번호":"💬 문구")+" 복사";},800);}' +
    'function markDone(r,b){google.script.run.withSuccessHandler(function(){document.getElementById("row_"+r).style.opacity="0.3"; b.innerText="확인됨"; b.disabled=true;}).setCheck(r);}' +
    '</script>';

  var header = '<div style="background:#e7f5ff; padding:15px; border-radius:10px; margin-bottom:20px; text-align:center;">' +
               '<b>📊 통합 건강 리포트 발송 (' + count + '건)</b><br>' +
               '<small>선택하신 범위 내 회원님들의 주간/월간 리포트입니다.</small></div>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput('<div style="font-family:sans-serif; background:#f8f9fa; padding:15px;">' + script + header + list + '</div>').setWidth(450).setHeight(700), "🏆 통합 건강 리포트");
}

/**
 * 밤 11시 30분경 자동으로 실행되어 그날의 최종 통계를 업무일지에 업데이트합니다.
 * (트리거 설정 필요: autoCloseDailyLog / 시간 기반 / 일일 타이머 / 오후 11시~자정)
 */
function autoCloseDailyLog() {
  try {
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. [자동 퇴실 처리] 미처 퇴실 처리가 안 된 회원들 강제 귀가 처리
    try {
      var resSheet = ss.getSheetByName("예약DB");
      var logSheet = ss.getSheetByName("출석기록");
      var resData = resSheet.getDataRange().getValues();
      var logData = logSheet.getDataRange().getValues();
      var forceTime = "22:00:00"; 
      
      for (var r = 1; r < resData.length; r++) {
        var rDateStr = (resData[r][3] instanceof Date) ? Utilities.formatDate(resData[r][3], "GMT+9", "yyyy-MM-dd") : String(resData[r][3]).split(" ")[0];
        if (rDateStr === dateStr && String(resData[r][9]).indexOf("입실") !== -1) {
          resSheet.getRange(r + 1, 10).setValue("귀가_자동마감");
          resSheet.getRange(r + 1, 12).setValue(forceTime);
        }
      }
      for (var l = 1; l < logData.length; l++) {
        var lDateStr = (logData[l][0] instanceof Date) ? Utilities.formatDate(logData[l][0], "GMT+9", "yyyy-MM-dd") : String(logData[l][0]).split(" ")[0];
        if (lDateStr === dateStr && (String(logData[l][11]).indexOf("입실") !== -1 || String(logData[l][11]) === "")) {
          logSheet.getRange(l + 1, 12).setValue("귀가(자동)");
          logSheet.getRange(l + 1, 13).setValue("22:00");
        }
      }
    } catch (e) { Logger.log("자동 퇴실 오류: " + e.toString()); }
    
    // 1. 최종 통계 가져오기
    var statsRes = getTodayWorkLogStats(dateStr);
    if (statsRes.error) {
      Logger.log("자동 마감 통계 조회 실패: " + statsRes.error);
      return;
    }
    var stats = statsRes.stats;
    
    // 2. 기존 일지가 있는지 확인
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var workLogSheet = ss.getSheetByName("업무일지");
    if (!workLogSheet) return;
    
    var wlData = workLogSheet.getDataRange().getValues();
    var targetRowIdx = -1;
    var existingLog = null;
    
    for (var i = 1; i < wlData.length; i++) {
      var rowDateRaw = wlData[i][0];
      var rowDateStr = (rowDateRaw instanceof Date) ? Utilities.formatDate(rowDateRaw, "GMT+9", "yyyy-MM-dd") : String(rowDateRaw).split(" ")[0];
      if (rowDateStr === dateStr) {
        targetRowIdx = i + 1;
        existingLog = wlData[i];
        break;
      }
    }
    
    // 3. 전송용 데이터 조립 (기존 내용 보존)
    var data = {
      date: dateStr,
      author: existingLog ? String(existingLog[1] || "") : "시스템 자동마감",
      jumpingList: existingLog ? String(existingLog[2] || "") : "(미기입)",
      muscleList: existingLog ? String(existingLog[3] || "") : "(미기입)",
      stats: stats,
      remarks: existingLog ? String(existingLog[13] || "") : "퇴근 전 미작성되어 시스템에 의해 자동 마감되었습니다.",
      issues: existingLog ? String(existingLog[14] || "") : ""
    };
    
    // 4. 저장 함수 호출
    var res = submitWorkLog(data);
    Logger.log(dateStr + " 자동 마감 결과: " + res.message);
    return res;
    
  } catch (e) {
    Logger.log("자동 마감 치명적 오류: " + e.toString());
    return { error: e.toString() };
  }
}

function setRowColor(r, sName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sName);
  if (sheet) sheet.getRange(r, 1, 1, sheet.getLastColumn()).setBackground("#dcfce7");
}

function getTargetSeasonLabelByInput(dateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var setupSheet = ss.getSheetByName("인바디 시즌제 등록");
    var setupData = setupSheet.getDataRange().getValues();
    var parts = dateStr.split("/");
    var targetDateOnly = new Date(2026, parseInt(parts[0]) - 1, parseInt(parts[1]));
    for (var i = 1; i < setupData.length; i++) {
      var seasonName = String(setupData[i][0] || "").trim();
      var awardDate = new Date(setupData[i][1]);
      if (!seasonName || isNaN(awardDate.getTime())) continue;
      var awardDateOnly = new Date(awardDate.getFullYear(), awardDate.getMonth(), awardDate.getDate());
      var diffDays = Math.round((awardDateOnly - targetDateOnly) / (1000 * 60 * 60 * 24));
      var weekNum = "";
      if (diffDays <= 1 && diffDays >= -5) weekNum = "3주차";
      else if (diffDays <= 8 && diffDays >= 2) weekNum = "2주차";
      else if (diffDays <= 15 && diffDays >= 9) weekNum = "1주차";
      if (weekNum !== "") return seasonName + " " + weekNum;
    }
    return "";
  } catch(e) { return ""; }
}

function backupRemainingSessions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("등록 현황");
  var targetSheet = ss.getSheetByName("잔여횟수 차감현황");
  if (!sourceSheet || !targetSheet) return;
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return;
  var sourceValues = sourceSheet.getRange(2, 1, lastRow - 1, sourceSheet.getLastColumn()).getValues();
  var dateStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm");
  var backupData = sourceValues.map(function(row) { return [dateStr].concat(row); });
  targetSheet.getRange(targetSheet.getLastRow() + 1, 1, backupData.length, backupData[0].length).setValues(backupData);
}

function runWeeklyStats() { generateStats('주간 통계'); }
function runMonthlyStats() { generateStats('월간 통계'); }

function generateStats(targetSheetName) {
  var ui = SpreadsheetApp.getUi();
  var startRes = ui.prompt('📅 [' + targetSheetName + '] 시작일', 'YYYY-MM-DD', ui.ButtonSet.OK_CANCEL);
  if (startRes.getSelectedButton() !== ui.Button.OK) return;
  var endRes = ui.prompt('📅 [' + targetSheetName + '] 종료일', 'YYYY-MM-DD', ui.ButtonSet.OK_CANCEL);
  if (endRes.getSelectedButton() !== ui.Button.OK) return;
  
  var startDate = new Date(startRes.getResponseText().trim());
  var endDate = new Date(endRes.getResponseText().trim());
  endDate.setHours(23, 59, 59, 999);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = ss.getSheetByName("등록 현황");
  var logSheet = ss.getSheetByName("출석기록");
  var memberSheet = ss.getSheetByName("회원명단"); 
  var targetSheet = ss.getSheetByName(targetSheetName);

  var regData = regSheet.getDataRange().getValues();
  var logData = logSheet.getDataRange().getValues();
  var memberData = memberSheet.getDataRange().getValues();
  var resultRows = []; var memberUpdateMap = {};

  for (var i = 1; i < regData.length; i++) {
    if (regData[i][10] === "진행중") {
      var mId = String(regData[i][2] || "").trim();
      var name = regData[i][3];
      var count = 0, tCount = 0, tMins = 0;
      for (var j = 1; j < logData.length; j++) {
        var logId = String(logData[j][3] || "").trim();
        var logTime = new Date(logData[j][1]);
        if (logId === mId && logTime >= startDate && logTime <= endDate) {
          count++; tCount += Number(logData[j][8] || 0); tMins += Number(logData[j][10] || 0);
        }
      }
      var formatted = Math.floor(tMins / 60) + "시간 " + (tMins % 60) + "분";
      resultRows.push([startDate.toLocaleDateString() + "~" + endDate.toLocaleDateString(), mId, name, count, tCount, formatted]);
      memberUpdateMap[mId] = {count: count, tims: tCount, timeStr: formatted};
    }
  }

  for (var k = 1; k < memberData.length; k++) {
    var mId = String(memberData[k][0] || "").trim();
    if (memberUpdateMap[mId]) {
      var ud = memberUpdateMap[mId];
      var col = (targetSheetName === '주간 통계') ? 17 : 20;
      memberSheet.getRange(k + 1, col, 1, 3).setValues([[ud.count, ud.tims, ud.timeStr]]);
    }
  }

  if (resultRows.length > 0) {
    targetSheet.insertRowsAfter(1, resultRows.length);
    targetSheet.getRange(2, 1, resultRows.length, 6).setValues(resultRows);
    ui.alert("✅ 통계 완료!");
  }
}

function run33ChallengeWeeklySettlement() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var sPrompt = ui.prompt('시즌 선택', '시즌명 입력', ui.ButtonSet.OK_CANCEL);
  if (sPrompt.getSelectedButton() != ui.Button.OK) return;
  var seasonName = sPrompt.getResponseText();

  var wPrompt = ui.prompt('주차 선택', '숫자만 (1,2,3)', ui.ButtonSet.OK_CANCEL);
  if (wPrompt.getSelectedButton() != ui.Button.OK) return;
  var weekNum = wPrompt.getResponseText();

  var setupSheet = ss.getSheetByName("인바디 시즌제 등록");
  var setupData = setupSheet.getDataRange().getValues();
  var awardDate = null;
  for (var i = 1; i < setupData.length; i++) {
    if (setupData[i][0] == seasonName) { awardDate = new Date(setupData[i][1]); break; }
  }
  if (!awardDate) return;

  var endDate = new Date(awardDate); endDate.setDate(awardDate.getDate() - ((3 - parseInt(weekNum)) * 7 + 1)); endDate.setHours(23, 59, 59);
  var startDate = new Date(endDate); startDate.setDate(endDate.getDate() - 2); 
  var startEndDate = new Date(awardDate); startEndDate.setDate(awardDate.getDate() - 22);
  var startStartDate = new Date(startEndDate); startStartDate.setDate(startEndDate.getDate() - 2);

  var inputSheet = ss.getSheetByName("인바디 입력");
  var regSheet = ss.getSheetByName("등록 현황");
  var targetSheet = ss.getSheetByName("33챌린지 주별 누적");
  var inputData = inputSheet.getDataRange().getValues();
  var regData = regSheet.getDataRange().getValues();
  var targetData = targetSheet.getDataRange().getValues();
  
  var weekXRows = [], week0Rows = [];

  for (var i = 1; i < regData.length; i++) {
    if (regData[i][10] == "진행중") {
      var mId = regData[i][2], name = regData[i][3];
      var curR = null, preR = null, baseR = null;

      for (var j = inputData.length - 1; j >= 1; j--) {
        var d = new Date(inputData[j][1]);
        if (inputData[j][3] == mId && d >= startDate && d <= endDate) { curR = inputData[j]; break; }
      }
      for (var j = inputData.length - 1; j >= 1; j--) {
        var d = new Date(inputData[j][1]);
        if (inputData[j][3] == mId && d >= startStartDate && d <= startEndDate) { baseR = inputData[j]; break; }
      }
      
      if (weekNum == "1") preR = baseR;
      else {
        var pLabel = seasonName + " " + (parseInt(weekNum) - 1) + "주차";
        for (var k = 0; k < targetData.length; k++) { if (targetData[k][1] == mId && targetData[k][0] == pLabel) { preR = targetData[k]; break; } }
      }

      if (!curR && !preR) continue;

      if (weekNum == "1" && baseR) {
        week0Rows.push([seasonName + " 0주차", mId, name, seasonName, baseR[4], baseR[5], baseR[6], (baseR[6]/baseR[4]*100).toFixed(2), 0, 0, 0, 0, baseR[7], "0주차 시작점"]);
      }

      var fW, fM, fF, note = "", wDW=0, fDW=0, fRW=0, fRS=0;
      if (!curR) {
        fW = preR[4]; fM = preR[5]; fF = preR[6]; note = "(미제출) 복사";
      } else {
        fW = curR[4]; fM = curR[5]; fF = curR[6];
      }

      if (preR) {
        wDW = (fW - preR[4]).toFixed(1); fDW = (fF - preR[6]).toFixed(1); fRW = (preR[6] != 0) ? ((fDW / preR[6]) * 100).toFixed(2) : 0;
      }
      if (baseR) fRS = (baseR[6] != 0) ? (((fF - baseR[6]) / baseR[6]) * 100).toFixed(2) : 0;

      weekXRows.push([seasonName + " " + weekNum + "주차", mId, name, seasonName, fW, fM, fF, (fW > 0 ? (fF/fW*100).toFixed(2) : 0), wDW, fDW, fRW, fRS, (curR ? curR[7] : ""), note]);
    }
  }

  var finalRows = weekXRows.concat(week0Rows);
  if (finalRows.length > 0) {
    targetSheet.insertRowsBefore(2, finalRows.length);
    targetSheet.getRange(2, 1, finalRows.length, 14).setValues(finalRows);
    var rankS = ss.getSheetByName("33챌린지 랭킹뷰");
    if (rankS) {
      rankS.getRange(2, 1, Math.max(1, rankS.getLastRow()), 14).clear();
      if (weekXRows.length > 0) rankS.getRange(2, 1, weekXRows.length, 14).setValues(weekXRows);
    }
    ui.alert("✅ 정산 완료");
  }
}

function showChallengeCard_Full() { createChallengeEntryCard(false); }
function showChallengeCard_Masked() { createChallengeEntryCard(true); }

function createChallengeEntryCard(isMasked) {
  var ui = SpreadsheetApp.getUi();
  var range = SpreadsheetApp.getActiveSpreadsheet().getActiveRange();
  var values = range.getDisplayValues();
  if (values.length < 1) return;

  var seasonRes = ui.prompt("🏆 시즌명", "입력", ui.ButtonSet.OK_CANCEL);
  var seasonName = seasonRes.getResponseText() || "33 챌린지";

  var memberNames = values.map(v => v[0].toString().trim()).filter(n => n !== "").sort();
  var finalNames = isMasked ? memberNames.map(n => n.length <= 1 ? n : n[0] + "*".repeat(n.length - 2) + n[n.length - 1]) : memberNames;

  var html = '<html><body style="font-family:sans-serif; text-align:center; padding:20px;">' +
             '<h2>🏆 ' + seasonName + ' 엔트리</h2>' +
             '<div style="display:grid; grid-template-columns:1fr 1fr; gap:10px;">' +
             finalNames.map(n => '<div style="padding:10px; background:#f8f9fa; border:1px solid #ddd;">' + n + '</div>').join('') +
           '</div>';
html += '</script></body></html>';
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(600), "챌린지 카드");
}
function getAdminReservationData(targetDateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("예약DB");
    if (!sheet) return [];
    
    // getDisplayValues()를 사용하여 시트에 보이는 글자 그대로 가져옴 (32분 오차 버그 방지)
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    
    // 헤더에서 동적으로 인덱스 찾기
    var idx = { name: 2, date: 3, time: 4, room: 7, people: 8, status: 9 }; 
    for (var h = 0; h < headers.length; h++) {
      var head = String(headers[h]).trim();
      if (head.indexOf("이름") !== -1) idx.name = h;
      if (head.indexOf("예약날짜") !== -1 || head.indexOf("날짜") !== -1) idx.date = h;
      if (head.indexOf("입실시간") !== -1 || head.indexOf("예약시간") !== -1) idx.time = h;
      if (head.indexOf("배정방") !== -1 || head.indexOf("방") !== -1) idx.room = h;
      if (head.indexOf("인원") !== -1) idx.people = h;
      if (head.indexOf("상태") !== -1) idx.status = h;
    }

    var results = [];
    for (var i = 1; i < data.length; i++) {
      var rDateStr = data[i][idx.date].split(" ")[0]; // 보이는 날짜 그대로
      
      if (rDateStr === targetDateStr) {
        var rawTime = data[i][idx.time];
        var timeStr = "시간미정";
        
        if (rawTime && rawTime.includes(':')) {
          var parts = rawTime.split(':');
          if (parts.length >= 2) {
            // HH:mm 형식으로 강제 고정 (9:0 -> 09:00)
            var hh = parts[0].trim().padStart(2, '0');
            var mm = parts[1].trim().padStart(2, '0');
            timeStr = hh + ":" + mm;
          }
        }
        
        // 최종 안전장치: 혹시라도 끝에 숫자가 아닌 찌꺼기가 남으면 제거
        timeStr = timeStr.replace(/[^0-9:]+$/, "");

        results.push({
          rowIdx: i + 1,
          time: timeStr, 
          name: data[i][idx.name],
          room: data[i][idx.room],
          people: data[i][idx.people],
          status: data[i][idx.status]
        });
      }
    }
    
    // 시간순 -> 방이름 순 정렬
    results.sort((a, b) => {
      if (a.time !== b.time) return a.time.localeCompare(b.time);
      return a.room.localeCompare(b.room);
    });
    
    return results;
  } catch (e) {
    return [];
  }
}

function deleteReservationByRow(rowIdx) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("예약DB");
    if (!sheet) return { error: "시트를 찾을 수 없습니다." };
    
    sheet.deleteRow(rowIdx);
    return { success: true };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * 특정 날짜의 벙개/휴무 설정 가져오기
 */
function getFlashSettingByDate(dateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("벙개테라피 및 휴일 설정");
    if (!sheet) return { found: false };
    
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rDate = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+9", "yyyy-MM-dd") : String(data[i][0]).split(" ")[0];
      if (rDate === dateStr) {
        return {
          found: true,
          type: data[i][1], // B열: 구분 (벙개/휴무)
          note: data[i][2]  // C열: 메모(벙개시간 등)
        };
      }
    }
    return { found: false };
  } catch (e) { return { error: e.toString() }; }
}

/**
 * 벙개/휴무 설정 저장 (덮어쓰기 포함)
 */
function saveFlashSetting(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("벙개테라피 및 휴일 설정");
    if (!sheet) {
      sheet = ss.insertSheet("벙개테라피 및 휴일 설정");
      sheet.appendRow(["날짜", "구분", "메모(벙개시간)"]);
      sheet.getRange("A1:C1").setBackground("#f3f3f3").setFontWeight("bold");
    }
    
    // 원장님 요청대로 헤더 강제 수정
    sheet.getRange("A1:C1").setValues([["날짜", "구분", "메모(벙개시간)"]]);
    
    var sheetData = sheet.getDataRange().getValues();
    var targetRow = -1;
    for (var i = 1; i < sheetData.length; i++) {
      if (!sheetData[i][0]) continue;
      var rDate = (sheetData[i][0] instanceof Date) ? Utilities.formatDate(sheetData[i][0], "GMT+9", "yyyy-MM-dd") : String(sheetData[i][0]).split(" ")[0];
      if (rDate === data.date) {
        targetRow = i + 1;
        break;
      }
    }
    
    if (targetRow !== -1) {
      // 기존 행 업데이트 (구분, 메모)
      sheet.getRange(targetRow, 2, 1, 2).setValues([[data.type, data.note]]);
    } else {
      // 새 행 추가
      sheet.appendRow([data.date, data.type, data.note]);
    }
    
    SpreadsheetApp.flush(); // 데이터 강제 동기화 (원장님 요청: 목록 즉시 갱신용)
    return { success: true };
  } catch (e) { return { error: e.toString() }; }
}

/**
 * 모든 벙개/휴무 설정 가져오기 (목록용)
 */
function getAllFlashSettings() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("벙개테라피 및 휴일 설정");
    if (!sheet) return [];
    
    // getDisplayValues()를 사용하여 날짜 형식을 안전하게 텍스트로 가져옴
    var data = sheet.getDataRange().getDisplayValues();
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var dStr = String(data[i][0]).trim();
      if (!dStr) continue;
      
      var type = String(data[i][1]).trim();
      // 기존 명칭과 새로운 명칭 모두 호환되도록 처리
      var normalizedType = (type === "벙개설정" || type === "벙개") ? "벙개" : "휴무";
      
      results.push({
        rowIdx: i + 1,
        date: dStr,
        type: normalizedType,
        note: data[i][2]
      });
    }
    // 최신순 정렬
    results.sort((a, b) => b.date.localeCompare(a.date));
    return results;
  } catch (e) { return []; }
}

/**
 * 벙개/휴무 설정 삭제
 */
function deleteFlashSettingByRow(rowIdx) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("벙개테라피 및 휴일 설정");
    if (!sheet) return { error: "시트를 찾을 수 없습니다." };
    sheet.deleteRow(rowIdx);
    return { success: true };
  } catch (e) { return { error: e.toString() }; }
}

/**
 * 밤 11시 30분경 자동으로 실행되어 그날의 최종 통계를 업무일지에 업데이트합니다.
 * (트리거 설정 필요: autoCloseDailyLog / 시간 기반 / 일일 타이머 / 오후 11시~자정)
 */
function autoCloseDailyLog() {
  try {
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    // 1. 최종 통계 가져오기
    var statsRes = getTodayWorkLogStats(dateStr);
    if (statsRes.error) {
      Logger.log("자동 마감 통계 조회 실패: " + statsRes.error);
      return;
    }
    var stats = statsRes.stats;
    
    // 2. 기존 일지가 있는지 확인
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var workLogSheet = ss.getSheetByName("업무일지");
    if (!workLogSheet) return;
    
    var wlData = workLogSheet.getDataRange().getValues();
    var targetRowIdx = -1;
    var existingLog = null;
    
    for (var i = 1; i < wlData.length; i++) {
      var rowDateRaw = wlData[i][0];
      var rowDateStr = (rowDateRaw instanceof Date) ? Utilities.formatDate(rowDateRaw, "GMT+9", "yyyy-MM-dd") : String(rowDateRaw).split(" ")[0];
      if (rowDateStr === dateStr) {
        targetRowIdx = i + 1;
        existingLog = wlData[i];
        break;
      }
    }
    
    // 3. 전송용 데이터 조립 (기존 내용 보존)
    var data = {
      date: dateStr,
      author: existingLog ? String(existingLog[1] || "") : "시스템 자동마감",
      jumpingList: existingLog ? String(existingLog[2] || "") : "(미기입)",
      muscleList: existingLog ? String(existingLog[3] || "") : "(미기입)",
      stats: stats,
      remarks: existingLog ? String(existingLog[13] || "") : "퇴근 전 미작성되어 시스템에 의해 자동 마감되었습니다.",
      issues: existingLog ? String(existingLog[14] || "") : ""
    };
    
    // 4. 저장 함수 호출
    var res = submitWorkLog(data);
    Logger.log(dateStr + " 자동 마감 결과: " + res.message);
    return res;
    
  } catch (e) {
    Logger.log("자동 마감 치명적 오류: " + e.toString());
    return { error: e.toString() };
  }
}
/**
 * [관리자 전용] 문자발송 대기 목록 가져오기
 */
function getSmsLogData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("문자발송");
    if (!sheet) return { error: "문자발송 시트가 없습니다." };
    
    var data = sheet.getDataRange().getDisplayValues();
    var results = [];
    
    // 헤더 제외, 역순으로 최신 것부터 확인
    for (var i = data.length - 1; i >= 1; i--) {
      var rowStatus = String(data[i][5] || "").trim();
      // 원장님 요청: 기본적으로 대기 상태만 반환하거나 클라이언트에서 처리하도록 전체 반환
      results.push({
        rowIdx: i + 1,
        time: data[i][0],
        name: data[i][1],
        phone: data[i][2],
        category: data[i][3],
        content: data[i][4],
        status: rowStatus
      });
    }
    return { success: true, logs: results };
  } catch (e) {
    return { error: "문자 목록 조회 실패: " + e.toString() };
  }
}

/**
 * 대시보드용 대기 중인 문자 개수 조회
 */
function getPendingSmsCount() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("문자발송");
    if (!sheet) return { count: 0 };
    
    var data = sheet.getDataRange().getDisplayValues();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][5]).trim() === "대기") count++;
    }
    return { count: count };
  } catch(e) { return { count: 0 }; }
}

/**
 * [자동화 전용] 매일 자정 혹은 주기적으로 실행하여 미방문/미등록 리스트를 최신화합니다.
 * '대기' 상태인 자동 생성 문자들을 삭제하고 다시 추출하여 방문 완료자를 명단에서 제외합니다.
 */
function autoRefreshSmsLists() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("문자발송");
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    // 아래에서부터 지워야 행 번호가 안 꼬임
    for (var i = data.length - 1; i >= 1; i--) {
      var category = String(data[i][3]);
      var status = String(data[i][5]).trim();
      
      // 자동 생성 카테고리이면서 아직 발송하지 않은(대기) 건만 상태 변경
      if ((category === "장기미방문" || category === "복귀권유") && status === "대기") {
        sheet.getRange(i + 1, 6).setValue("완료(재방문)");
      }
    }
    
    // 다시 추출 (이미 방문한 사람은 lastAttendanceMap에 의해 제외됨)
    checkLongTermAbsentees();
    checkInactiveMembers();
    
    Logger.log("SMS 발송 대기 목록 자동 최신화 완료");
  } catch (e) {
    Logger.log("SMS 자동 최신화 오류: " + e.toString());
  }
}

/**
 * [관리자 전용] 문자발송 상태 업데이트 (완료 등)
 */
function updateSmsStatus(rowIdx, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("문자발송");
    if (!sheet) return { error: "시트를 찾을 수 없습니다." };
    
    sheet.getRange(rowIdx, 6).setValue(status);
    return { success: true, message: "상태가 " + status + "(으)로 업데이트되었습니다." };
  } catch (e) {
    return { error: "상태 업데이트 실패: " + e.toString() };
  }
}

/**
 * [관리자 전용] 7일 이상 미방문 회원 추출 및 문자 생성
 */
function checkLongTermAbsentees() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황");
    var logSheet = ss.getSheetByName("출석기록");
    var smsSheet = ss.getSheetByName("문자발송");
    
    if (!regSheet || !logSheet || !smsSheet) return { error: "필요한 시트(등록현황/출석기록/문자발송)가 없습니다." };
    
    var regData = regSheet.getDataRange().getDisplayValues();
    var logData = logSheet.getDataRange().getDisplayValues();
    var smsData = smsSheet.getDataRange().getDisplayValues();
    
    var regCols = getRegColumnIndices(regSheet);
    var logCols = getAttendanceColumnIndices(logSheet);
    
    var now = new Date();
    var sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(now.getDate() - 7);
    
    // 1. 마지막 출석일 맵 작성 (폰번호 기준)
    var lastAttendanceMap = {};
    for (var i = 1; i < logData.length; i++) {
      var phone = String(logData[i][logCols.phone] || "").replace(/[^0-9]/g, "");
      var dateStr = logData[i][logCols.date];
      if (!phone || !dateStr) continue;
      
      var attendanceDate = new Date(dateStr);
      if (!lastAttendanceMap[phone] || attendanceDate > lastAttendanceMap[phone]) {
        lastAttendanceMap[phone] = attendanceDate;
      }
    }
    
    // 2. 이미 발송 대기 중인 목록 확인 (중복 생성 방지)
    var existingSmsMap = {};
    for (var j = 1; j < smsData.length; j++) {
      var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
      var sCategory = smsData[j][3];
      var sStatus = smsData[j][5];
      if (sCategory === "장기미방문" && (sStatus === "대기" || sStatus === "완료")) {
        existingSmsMap[sPhone] = true;
      }
    }
    
    // 3. 미방문자 추출 및 문자 생성
    var count = 0;
    var addedNames = [];
    
    for (var k = 1; k < regData.length; k++) {
      var status = String(regData[k][regCols.status]).trim();
      if (status !== "진행중" && status !== "진행 중") continue;
      
      var phone = String(regData[k][regCols.phone] || "").replace(/[^0-9]/g, "");
      var name = regData[k][regCols.name];
      if (!phone) continue;
      
      var lastDate = lastAttendanceMap[phone];
      
      // 마지막 출석일이 없거나(한번도 안옴), 7일 이상 지난 경우
      if (!lastDate || lastDate < sevenDaysAgo) {
        // 이미 보냈거나 대기 중이면 패스
        if (existingSmsMap[phone]) continue;
        
        var cleanName = name.replace(/\d{4}$/, ""); // 이름 뒤 번호 제거
        var msg = cleanName + "회원님, 노형 점핑클럽입니다. 😊 클럽에서 뵙지 못한 지 너무 오래된 것 같아요. 😢 혹시 어디 불편하신 건 아니시죠? 바쁜 일상 속에서도 건강만큼은 꼭 챙기셨으면 하는 마음에 연락드렸습니다. ❤️ 잠시 짬을 내어 신나게 뛰러 오시는 건 어때요? 🏃‍♀️ 아니면 편안하게 테라피를 하러 오셔도 좋고요. 🔥 이번 주는 꼭 얼굴 뵀으면 좋겠어요! 😊";
        
        // 휴대폰 번호 보정 (0 누락 방지)
        var formattedPhone = formatPhoneForSms(regData[k][regCols.phone]);

        // 문자발송 시트에 추가
        smsSheet.appendRow([
          Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm"),
          cleanName,
          formattedPhone,
          "장기미방문",
          msg,
          "대기"
        ]);
        
        count++;
        addedNames.push(cleanName);
      }
    }
    
    return { 
      success: true, 
      count: count, 
      message: count > 0 ? count + "명의 미방문 회원을 추출했습니다." : "새로운 미방문 회원이 없습니다.",
      addedNames: addedNames
    };
    
  } catch (e) {
    return { error: "미방문자 추출 중 오류: " + e.toString() };
  }
}

/**
 * [관리자 전용] 등록이 끊긴 지 14일 이상 된 장기 미등록 회원 추출
 */
function checkInactiveMembers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황");
    var smsSheet = ss.getSheetByName("문자발송");
    
    if (!regSheet || !smsSheet) return { error: "필요한 시트(등록현황/문자발송)가 없습니다." };
    
    var regData = regSheet.getDataRange().getDisplayValues();
    var smsData = smsSheet.getDataRange().getDisplayValues();
    var regCols = getRegColumnIndices(regSheet);
    
    var now = new Date();
    var fourteenDaysAgo = new Date();
    fourteenDaysAgo.setDate(now.getDate() - 14);
    
    // 1. 회원별 상태 체크 (가장 최근 만료일 및 활성 여부)
    var memberInfoMap = {};
    for (var i = 1; i < regData.length; i++) {
      var phone = String(regData[i][regCols.phone] || "").replace(/[^0-9]/g, "");
      var status = String(regData[i][regCols.status] || "").trim(); 
      var expireDateStr = regData[i][regCols.expire];
      var name = regData[i][regCols.name];
      
      if (!phone) continue;
      
      if (!memberInfoMap[phone]) {
        memberInfoMap[phone] = { isActive: false, lastExpire: new Date(0), name: name, phoneRaw: regData[i][regCols.phone] };
      }
      
      // 하나라도 진행중이면 활성 회원
      if (status === "진행중" || status === "진행 중") {
        memberInfoMap[phone].isActive = true;
      }
      
      // 만료일 업데이트
      if (expireDateStr) {
        var expDate = new Date(expireDateStr);
        if (expDate > memberInfoMap[phone].lastExpire) {
          memberInfoMap[phone].lastExpire = expDate;
        }
      }
    }
    
    // 2. 이미 발송 목록에 있는지 확인
    var existingSmsMap = {};
    for (var j = 1; j < smsData.length; j++) {
      var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
      var sCategory = smsData[j][3];
      var sStatus = smsData[j][5];
      // 이미 복귀권유 대기 중이거나 최근 발송 완료된 경우 제외
      if (sCategory === "복귀권유" && (sStatus === "대기" || sStatus === "완료")) {
        existingSmsMap[sPhone] = true;
      }
    }
    
    // 3. 비활성 회원 중 14일 이상 된 분들 추출
    var count = 0;
    var addedNames = [];
    
    var phones = Object.keys(memberInfoMap);
    for (var pIdx = 0; pIdx < phones.length; pIdx++) {
      var m = memberInfoMap[phones[pIdx]];
      
      // 활성 회원이 아니고, 마지막 만료일이 14일보다 더 과거라면
      if (!m.isActive && m.lastExpire < fourteenDaysAgo && m.lastExpire.getTime() > 0) {
        if (existingSmsMap[phones[pIdx]]) continue;
        
        var cleanName = m.name.replace(/\d{4}$/, "");
        var msg = "";
        
        // 잔여 횟수 확인 (memberInfoMap 생성 시 합산 로직 추가 필요 - 아래 루프에서 보완)
        var totalRemain = 0;
        for (var k = 1; k < regData.length; k++) {
          if (String(regData[k][regCols.phone] || "").replace(/[^0-9]/g, "") === phones[pIdx]) {
            var r = parseInt(regData[k][regCols.remain]) || 0;
            totalRemain += r;
          }
        }

        if (totalRemain >= 5) {
          // 유형 A: 횟수가 5회 이상 남은 분들 (보너스 제안)
          msg = cleanName + "회원님! 노형점핑 클럽입니다. 😊 잘 지내고 계시죠? 지난번 이용하실 때 아깝게 남은 횟수들이 기간이 지나 만료되는 바람에 저희도 참 마음이 쓰였답니다. 😢 이번 기회에 다시 오시면 보너스권으로 넉넉히 채워드릴게요. 🔥 다시 한번 활기차게 건강관리 시작해 보시는 건 어떨까요? 기다리고 있겠습니다. ❤️";
        } else {
          // 유형 B: 다 썼거나 5회 미만 남은 분들 (열정 재점화)
          msg = cleanName + "회원님, 노형점핑클럽입니다. 😊 그동안 잘 지내셨나요? 문득 회원님과 함께 신나게 땀 흘리며 운동하던 시간이 떠올라 소식 전해요. 🏃‍♀️ 다시 운동을 시작하고 싶은 마음이 드신다면 언제든 가볍게 들러주세요. 제가 그동안 더 연마한(?) 비법으로 점핑 자세 하나하나 기깔나게 다시 잡아드릴게요. 🔥 우리 다시 한번 뜨거운 열정을 불태워봐요! 😊";
        }
        
        // 휴대폰 번호 보정 (0 누락 방지)
        var formattedPhone = formatPhoneForSms(m.phoneRaw);

        smsSheet.appendRow([
          Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm"),
          cleanName,
          formattedPhone,
          "복귀권유",
          msg,
          "대기"
        ]);
        
        count++;
        addedNames.push(cleanName);
      }
    }
    
    return { 
      success: true, 
      count: count, 
      message: count > 0 ? count + "명의 복귀 권유 대상을 추출했습니다." : "새로운 복귀 권유 대상이 없습니다.",
      addedNames: addedNames
    };
    
  } catch (e) {
    return { error: "복귀 대상 추출 중 오류: " + e.toString() };
  }
}

/**
 * [공용] 휴대폰 번호 정규화 (010-0000-0000 형식 강제)
 */
function formatPhoneNumber(phoneStr) {
  if (!phoneStr) return "";
  var phone = String(phoneStr).replace(/[^0-9]/g, "");
  
  // 1012345678 -> 01012345678 (맨 앞 0이 빠진 경우 보정)
  if (phone.startsWith("1") && (phone.length === 10 || phone.length === 9)) {
    phone = "0" + phone;
  }
  
  if (phone.length === 11) {
    return phone.replace(/(\d{3})(\d{4})(\d{4})/, "$1-$2-$3");
  } else if (phone.length === 10) {
    return phone.replace(/(\d{3})(\d{3})(\d{4})/, "$1-$2-$3");
  }
  
  return phoneStr; 
}

/**
 * [관리자용] 시트 내의 모든 잘못된 전화번호 일괄 수정
 * 실행 방법: 웹앱 URL 뒤에 ?action=fixAllPhoneNumbersInSheet 붙여서 접속
 */
function fixAllPhoneNumbersInSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ["회원명단", "출석기록", "등록 현황", "예약DB", "문자발송"];
  var fixedCount = 0;
  
  sheets.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    var headers = data[0];
    var phoneCols = [];
    
    // "폰", "번호", "연락처", "휴대폰", "ID" 등이 들어간 모든 열 찾기
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]);
      if (h.indexOf("폰") !== -1 || h.indexOf("번호") !== -1 || h.indexOf("연락처") !== -1 || h.indexOf("휴대폰") !== -1 || h.indexOf("ID") !== -1) {
        phoneCols.push(i);
      }
    }
    
    if (phoneCols.length > 0) {
      phoneCols.forEach(function(colIdx) {
        for (var r = 1; r < data.length; r++) {
          var original = String(data[r][colIdx] || "");
          if (!original || original.trim() === "" || original === "-") continue;
          
          var formatted = formatPhoneNumber(original);
          
          // [안전장치] 정말 전화번호인 경우에만 업데이트 (숫자 9~10자리 & 1로 시작)
          var digitsOnly = original.replace(/[^0-9]/g, "");
          if (original !== formatted && (digitsOnly.length === 9 || digitsOnly.length === 10) && digitsOnly.startsWith("1")) {
            sheet.getRange(r + 1, colIdx + 1).setValue("'" + formatted); 
            fixedCount++;
          }
        }
      });
    }
  });
  
  return { success: true, message: "총 " + fixedCount + "개의 전화번호를 정규화(010-xxxx-xxxx)했습니다." };
}

function formatPhoneForSms(phoneStr) {
  return formatPhoneNumber(phoneStr);
}

/**
 * [추가] 테라피 예약 시 사용 가능한 회원권이나 보너스권이 있는지 검사하는 함수
 */
function checkMemberTicket(name, phone) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    var memberSheet = ss.getSheetByName("회원명단");
    
    if (!regSheet) return { success: false, error: "등록 현황 시트를 찾을 수 없습니다." };
    
    var phoneClean = String(phone || "").replace(/[^0-9]/g, "");
    var regData = regSheet.getDataRange().getValues();
    var cols = getRegColumnIndices(regSheet);
    
    var hasValidTicket = false;
    var foundTickets = [];
    
    // 1. 등록현황 확인 (테라피, 점핑회수권, 월권 등)
    for (var i = 1; i < regData.length; i++) {
      var rPhone = String(regData[i][cols.phone]).replace(/[^0-9]/g, "");
      var rStatus = String(regData[i][cols.status]).trim();
      
      if (rPhone === phoneClean && (rStatus === "진행중" || rStatus === "진행 중")) {
        var membership = String(regData[i][cols.membership]);
        var remainRaw = regData[i][cols.remain];
        var remain = parseInt(remainRaw) || 0;
        
        // 테라피권, 점핑 회수권, 월권 등 예약 가능한 모든 권종 포함
        if (membership.indexOf("테라피") !== -1 || membership.indexOf("점핑") !== -1 || membership.indexOf("회") !== -1 || membership.indexOf("월권") !== -1) {
          if (remain > 0 || String(remainRaw).indexOf("무제한") !== -1) {
            hasValidTicket = true;
            foundTickets.push(membership + "(" + remainRaw + "회)");
          }
        }
      }
    }
    
    // 2. 보너스권 확인 (회원명단 J열 -> Index 9)
    if (memberSheet) {
      var mData = memberSheet.getDataRange().getValues();
      for (var j = 1; j < mData.length; j++) {
        var mPhone = String(mData[j][2]).replace(/[^0-9]/g, "");
        if (mPhone === phoneClean) {
          var bonus = parseInt(mData[j][9]) || 0;
          if (bonus > 0) {
            hasValidTicket = true;
            foundTickets.push("보너스권(" + bonus + "회)");
          }
          break;
        }
      }
    }
    
    return {
      success: true,
      hasValidTicket: hasValidTicket,
      tickets: foundTickets.join(", ") || "없음"
    };
  } catch (e) {
    return { success: false, error: "서버 오류: " + e.toString() };
  }
}


// ──────────────────────────────────────────────
// 33 챌린지 게임 시스템 - 기초 셋업
// ──────────────────────────────────────────────

/**
 * 33 챌린지에 필요한 4개 시트를 자동으로 생성하고 헤더를 설정합니다.
 * 앱스 스크립트 에디터에서 이 함수를 선택하고 [실행]을 눌러주세요.
 */
function setup33ChallengeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = {
    "33챌린지_설정": ["시즌명", "시작일", "종료일", "상태", "참여인원", "설명"],
    "33챌린지_기록": ["타임스탬프", "회원명", "연락처", "구분", "항목", "내용", "점수", "이미지URL", "상태"],
    "33챌린지_인바디": ["날짜", "회원명", "연락처", "체중", "골격근량", "체지방률", "변화점수", "비고"],
    "33챌린지_점수": ["회원명", "연락처", "총경험치", "퀘스트점수", "습관점수", "인바디점수", "현재레벨", "최종업데이트"]
  };

  for (var name in sheets) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(sheets[name]);
      
      // 디자인 세팅 (첫 줄 고정 및 강조)
      sheet.setFrozenRows(1);
      sheet.getRange("1:1").setBackground("#f3f3f3").setFontWeight("bold").setHorizontalAlignment("center");
      
      Logger.log("✅ 시트 생성 완료: " + name);
    } else {
      Logger.log("ℹ️ 이미 존재하는 시트입니다: " + name);
    }
  }
  
  SpreadsheetApp.getUi().alert("33 챌린지 기초 시트 4개가 성공적으로 구축되었습니다!");
}

/**
 * 회원의 33 챌린지 현재 상태(레벨, 점수, 오늘 퀘스트 등)를 조회합니다.
 */
function get33ChallengeStatus(phone) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var scoreSheet = ss.getSheetByName("33챌린지_점수");
    var recordSheet = ss.getSheetByName("33챌린지_기록");
    var settingSheet = ss.getSheetByName("33챌린지_설정");
    
    var phoneOnly = String(phone).replace(/[^0-9]/g, "");
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    // 1. 현재 진행 중인 시즌 정보 가져오기
    var settings = settingSheet.getDataRange().getDisplayValues();
    var activeSeason = "33 챌린지"; // 기본값
    for (var i = 1; i < settings.length; i++) {
      if (settings[i][3] === "진행") {
        activeSeason = settings[i][0];
        break;
      }
    }
    
    // 2. 회원의 현재 점수 및 레벨 조회
    var scoreData = scoreSheet.getDataRange().getDisplayValues();
    var myScore = { level: 1, totalExp: 0, questExp: 0, habitExp: 0, inbodyExp: 0 };
    var found = false;
    
    for (var j = 1; j < scoreData.length; j++) {
      if (scoreData[j][1].replace(/[^0-9]/g, "") === phoneOnly) {
        myScore = {
          level: Number(scoreData[j][6] || 1),
          totalExp: Number(scoreData[j][2] || 0),
          questExp: Number(scoreData[j][3] || 0),
          habitExp: Number(scoreData[j][4] || 0),
          inbodyExp: Number(scoreData[j][5] || 0)
        };
        found = true;
        break;
      }
    }
    
    // 3. 오늘의 퀘스트 완료 현황 조회
    var records = recordSheet.getDataRange().getDisplayValues();
    var todayQuests = {
      water: 0,
      habits: [],
      photos: []
    };
    
    for (var k = 1; k < records.length; k++) {
      var rDate = records[k][0].split(" ")[0]; // 타임스탬프에서 날짜만 추출
      var rPhone = records[k][2].replace(/[^0-9]/g, "");
      
      if (rDate === todayStr && rPhone === phoneOnly) {
        var type = records[k][3]; // 구분 (물/식단/습관/인증)
        var item = records[k][4]; // 항목명
        
        if (type === "물") todayQuests.water = Number(records[k][5]);
        else if (type === "습관") todayQuests.habits.push(item);
        else if (type === "식단" || type === "인증") todayQuests.photos.push(item);
      }
    }
    
    return {
      success: true,
      season: activeSeason,
      status: myScore,
      today: todayQuests,
      isRegistered: found
    };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 회원의 활동(물 마시기, 습관 체크 등)을 기록하고 점수를 부여합니다.
 */
function submit33Action(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var recordSheet = ss.getSheetByName("33챌린지_기록");
    var scoreSheet = ss.getSheetByName("33챌린지_점수");
    
    var now = new Date();
    var timestamp = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm:ss");
    var phoneOnly = String(data.phone).replace(/[^0-9]/g, "");
    
    // 1. 기록 추가 (appendRow)
    // ["타임스탬프", "회원명", "연락처", "구분", "항목", "내용", "점수", "이미지URL", "상태"]
    recordSheet.appendRow([
      timestamp, 
      data.name, 
      data.phone, 
      data.type, 
      data.item, 
      data.value || "", 
      data.score || 0, 
      data.imageUrl || "",
      (data.type === "식단" || data.type === "퀘스트") ? "대기" : "승인"
    ]);
    
    // 2. 점수 합산 (Upsert 로직)
    var scoreData = scoreSheet.getDataRange().getDisplayValues();
    var foundIdx = -1;
    for (var i = 1; i < scoreData.length; i++) {
      if (scoreData[i][1].replace(/[^0-9]/g, "") === phoneOnly) {
        foundIdx = i + 1;
        break;
      }
    }
    
    var scoreToAdd = Number(data.score || 0);
    var questAdd = (data.type === "퀘스트" || data.type === "식단") ? scoreToAdd : 0;
    var habitAdd = (data.type === "습관" || data.type === "물") ? scoreToAdd : 0;

    if (foundIdx > 0) {
      // 기존 회원 업데이트
      var currentTotal = Number(scoreSheet.getRange(foundIdx, 3).getValue() || 0);
      var currentQuest = Number(scoreSheet.getRange(foundIdx, 4).getValue() || 0);
      var currentHabit = Number(scoreSheet.getRange(foundIdx, 5).getValue() || 0);
      
      scoreSheet.getRange(foundIdx, 3).setValue(currentTotal + scoreToAdd);
      scoreSheet.getRange(foundIdx, 4).setValue(currentQuest + questAdd);
      scoreSheet.getRange(foundIdx, 5).setValue(currentHabit + habitAdd);
      scoreSheet.getRange(foundIdx, 8).setValue(timestamp); // 최종업데이트
      
      // 레벨 계산 (예: 100점당 1레벨)
      var newLevel = Math.floor((currentTotal + scoreToAdd) / 100) + 1;
      scoreSheet.getRange(foundIdx, 7).setValue(newLevel);
    } else {
      // 신규 챌린저 등록
      scoreSheet.appendRow([
        data.name, 
        data.phone, 
        scoreToAdd, 
        questAdd, 
        habitAdd, 
        0, 1, timestamp
      ]);
    }
    
    return { success: true, totalExp: (currentTotal || 0) + scoreToAdd };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 📜 전령의 기둥 (실시간 공지) 관련 로직
 */
function getPillarNotice() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("마을_공지") || ss.insertSheet("마을_공지");
    if (sheet.getLastRow() < 1) return { content: "오늘도 건강한 하루 되세요! 지니 월드에 오신 것을 환영합니다." };
    
    var data = sheet.getDataRange().getValues();
    // 가장 마지막(최신) 활성화된 공지 가져오기
    for (var i = data.length - 1; i >= 1; i--) {
      if (data[i][3] === true || data[i][3] === "TRUE") {
        return { content: data[i][1], category: data[i][2] };
      }
    }
    return { content: "특별한 계시가 없는 평화로운 날입니다." };
  } catch(e) { return { content: "전령의 기둥이 흔들리고 있습니다..." }; }
}

function updatePillarNotice(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("마을_공지") || ss.insertSheet("마을_공지");
    
    // 기존 공지 비활성화 (선택 사항)
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // 모든 공지 비활성화 처리 로직 생략 가능 (최신순으로 가져오므로)
    }
    
    sheet.appendRow([new Date(), data.content, data.category || "일반", true]);
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

/**
 * ✍️ 지혜의 보물고 (집필실) 관련 로직
 */
function getWisdomTips() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("지혜의_보물고") || ss.insertSheet("지혜의_보물고");
    if (sheet.getLastRow() < 1) return [];
    
    var data = sheet.getDataRange().getDisplayValues();
    var tips = [];
    for (var i = 1; i < data.length; i++) {
      tips.push({
        id: i,
        date: data[i][0],
        title: data[i][1],
        content: data[i][2],
        category: data[i][3],
        author: data[i][4]
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
      sheet.appendRow(["날짜", "제목", "내용", "카테고리", "작성자", "조회수"]);
    }
    sheet.appendRow([new Date(), data.title, data.content, data.category, "길드마스터", 0]);
    return { success: true };
  } catch(e) { return { success: false, error: e.toString() }; }
}

/**
 * ✨ 이장의 축복 (추가 보상) 관련 로직
 */
function getRecentCertifications() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지_기록");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getDisplayValues();
    var recent = [];
    for (var i = 1; i < data.length; i++) {
      // 오직 '대기' 상태인 식단/퀘스트만 추출
      if ((data[i][3] === "식단" || data[i][3] === "퀘스트") && data[i][8] === "대기") {
        recent.push({
          rowIdx: i + 1,
          date: data[i][0],
          name: data[i][1],
          phone: data[i][2],
          type: data[i][3],
          item: data[i][4],
          content: data[i][5],
          imageUrl: data[i][7]
        });
      }
    }
    return recent;
  } catch(e) { return []; }
}

function blessAction(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지_기록");
    var rowIdx = data.rowIdx;
    
    // 1. 해당 행의 상태를 '승인'으로 변경 (9번째 열: I열)
    sheet.getRange(rowIdx, 9).setValue("승인");
    
    // 2. 승인 시 보너스 점수(+5) 하사
    var result = submit33Action({
      phone: data.phone,
      name: data.name,
      type: "축복",
      item: "이장의축복",
      value: "특별 보너스",
      score: 5
    });
    
    // 3. (옵션) 인증샷 자체의 기본 점수도 이때 합산하고 싶다면 여기서 추가 로직 구현 가능
    // 현재 구조는 인증샷 등록 시 0점(대기)이었다가 승인 시 점수가 들어가는 방식이 더 권위적임
    
    return result;
  } catch(e) { return { success: false, error: e.toString() }; }
}

/**
 * 📱 앱 접속 기록 및 일일 활성 모험가 집계
 */
function recordAppAccess(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("앱_접속_로그") || ss.insertSheet("앱_접속_로그");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["날짜", "시간", "연락처", "이름"]);
    }
    var now = new Date();
    sheet.appendRow([
      Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd"),
      Utilities.formatDate(now, "GMT+9", "HH:mm:ss"),
      data.phone,
      data.name
    ]);
    return { success: true };
  } catch(e) { return { success: false }; }
}

function getDailyActiveAdventurers() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("앱_접속_로그");
    if (!sheet) return 0;
    
    var data = sheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    var activeUsers = new Set();
    
    for (var i = 1; i < data.length; i++) {
      var logDate = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+9", "yyyy-MM-dd") : String(data[i][0]);
      if (logDate === todayStr) {
        activeUsers.add(data[i][2]); // 연락처(Unique ID) 기준
      }
    }
    return activeUsers.size;
  } catch(e) { return 0; }
}

/**
 * ⚡ 돌발 퀘스트 및 이벤트 관리 엔진
 */
function getActiveEvents() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("마을_이벤트") || ss.insertSheet("마을_이벤트");
    if (sheet.getLastRow() < 1) return [];
    
    var data = sheet.getDataRange().getValues();
    var now = new Date();
    var activeEvents = [];
    
    for (var i = 1; i < data.length; i++) {
      var endTime = new Date(data[i][3]);
      var status = data[i][4];
      
      // 기한이 남아있고 활성 상태인 이벤트만 추출
      if (endTime > now && status === "ACTIVE") {
        activeEvents.push({
          id: i,
          title: data[i][0],
          description: data[i][1],
          multiplier: data[i][2] || 1,
          endTime: Utilities.formatDate(endTime, "GMT+9", "HH:mm"),
          type: data[i][5] // 'BOOST' or 'NEW'
        });
      }
    }
    return activeEvents;
  } catch(e) { return []; }
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
