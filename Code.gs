/**
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * v1.0 - Core Routing & Database Setup
 */

function getArchiveFeed(payload) {
  try {
    var p = 1;
    var limit = 12;
    if (payload) {
      if (payload.page) p = parseInt(payload.page) || 1;
      if (payload.limit) limit = parseInt(payload.limit) || 12;
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("아카이브");
    if (!sheet) return { success: false, error: "'아카이브' 시트를 찾을 수 없습니다." };
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, data: { photos: [], total: 0 } };
    
    // 💡 리액션 정보가 있는 10번째 열(J열)까지 데이터를 가져옵니다.
    var data = sheet.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
    var photos = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var photoId = String(row[7] || "").trim();
      
      // 사진 ID가 실제로 존재하는 유효한 자랑 이미지들만 필터링
      if (photoId && photoId.length > 5) {
        var imageUrl = "https://lh3.googleusercontent.com/d/" + photoId;
        
        // 💡 10번째 열(Index 9)에서 리액션 JSON을 파싱합니다.
        var reactions = { cool: [], best: [], cheer: [] };
        var reactionJson = String(row[9] || "{}").trim();
        try { reactions = JSON.parse(reactionJson); } catch(e) {}
        
        photos.push({
          rowIdx: i + 2, // 💡 실제 시트에서의 행 번호 (2행부터 시작하므로 i + 2)
          url: imageUrl,
          description: String(row[6] || "소중한 인증의 한 장면 🌟"),
          author: String(row[2] || "모험가"),
          date: String(row[0] || ""),
          time: String(row[1] || ""),
          type: String(row[4] || "기록"),
          item: String(row[5] || ""),
          score: String(row[8] || "0"),
          reactions: reactions // 💡 리액션 데이터 포함!
        });
      }
    }
    
    // 최근 등록 사진이 위로 오도록 정렬
    photos.reverse();
    
    var total = photos.length;
    var startIndex = (p - 1) * limit;
    var slicedPhotos = photos.slice(startIndex, startIndex + limit);
    
    return {
      success: true,
      data: {
        photos: slicedPhotos,
        total: total
      }
    };
  } catch (e) {
    return { success: false, error: "사진 피드 로딩 실패: " + e.toString() };
  }
}

/**
 * [아카이브] 진행 중인 돌발 퀘스트 가져오기
 */
/**
 * [아카이브] 폴라로이드 리액션 업데이트
 */
function updateReaction(rowIdx, reactionType, phone) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("아카이브");
    if (!sheet) return { success: false };
    
    var lock = LockService.getScriptLock();
    lock.waitLock(5000);
    
    var reactionJson = String(sheet.getRange(rowIdx, 10).getValue() || "{}");
    var reactions = { cool: [], best: [], cheer: [] };
    try { reactions = JSON.parse(reactionJson); } catch(e) {}
    
    if (!reactions[reactionType]) reactions[reactionType] = [];
    
    var list = reactions[reactionType];
    var idx = list.indexOf(phone);
    if (idx > -1) {
      list.splice(idx, 1);
    } else {
      list.push(phone);
    }
    
    sheet.getRange(rowIdx, 10).setValue(JSON.stringify(reactions));
    lock.releaseLock();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * [아카이브] 진행 중인 돌발 퀘스트 가져오기
 */
function getActiveEvents() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndCreateQuestRegistrySheet();
    var data = sheet.getDataRange().getDisplayValues();
    
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var events = [];
    
    for (var i = 1; i < data.length; i++) {
      var type = data[i][2]; // C열: 유형 ("이장" or "글리코겐")
      var dateStr = data[i][1]; // B열: 시행일 (yyyy-MM-dd)
      var title = data[i][3]; // D열: 퀘스트명
      
      if (type === "이장" && dateStr === todayStr) {
        events.push({ title: title });
      }
    }
    return events;
  } catch (e) { 
    return []; 
  }
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

    // [v44.188] 신원 정보 최종 검증
    if (!payload.name || !payload.phone || payload.name === "무명 모험가") {
      return { success: false, error: "회원 정보가 누락되었습니다. 다시 로그인해 주세요." };
    }

    // [v44.188] 전화번호 정규화 및 중복 인증 방지 필터링
    var formattedPhone = formatPhoneNumber(payload.phone);
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, "GMT+9", "HH:mm:ss");

    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var arData = sheet.getRange(Math.max(2, lastRow - 100), 1, Math.min(lastRow - 1, 101), 9).getValues();
      var isDuplicate = false;
      var cleanPhone = formattedPhone.replace(/[^0-9]/g, "");
      for (var i = arData.length - 1; i >= 0; i--) {
        var rowDate = (arData[i][0] instanceof Date) ? Utilities.formatDate(arData[i][0], "GMT+9", "yyyy-MM-dd") : String(arData[i][0]).split(" ")[0];
        if (rowDate === dateStr) {
          var rowPhone = String(arData[i][3]).replace(/[^0-9]/g, "");
          var rowItem = String(arData[i][5]);
          if (rowPhone === cleanPhone && rowItem === payload.item) {
            isDuplicate = true;
            break;
          }
        }
      }
      if (isDuplicate) {
        Logger.log("⚠️ [중복인증 감지] " + payload.name + " (" + formattedPhone + ") - " + payload.item + " 이미 오늘 인증되었습니다.");
        return { 
          success: true, 
          message: "이미 오늘 인증이 완료되었습니다.",
          debugInfo: "DUP_PREVENTED",
          serverVersion: "v44.203_optimized"
        };
      }
    }

    var photoId = "";
    var photoError = "";

    // [v44.196] 데이터 수송 정밀 진단 로그
    if (payload.image) {
      Logger.log("[v44.196] 사진 데이터 수신됨! 길이: " + payload.image.length);
    } else {
      Logger.log("[v44.196] 경보! 사진 데이터가 누락되어 도착함.");
    }
    
    // [v44.199] 사진 저장 로직 (속도 향상을 위한 폴더 ID 캐싱 적용)
    if (payload.image && String(payload.image).length > 100) {
      try {
        var folderName = "GenieWorld_Archive";
        var scriptProps = PropertiesService.getScriptProperties();
        var cachedFolderId = scriptProps.getProperty("archive_folder_id");
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
          folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
          folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          scriptProps.setProperty("archive_folder_id", folder.getId());
        }
        
        var base64Data = payload.image;
        if (base64Data.indexOf(",") > -1) {
          base64Data = base64Data.split(",")[1];
        }
        base64Data = base64Data.replace(/\s/g, '');
        
        var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), "image/jpeg", (payload.name || "user") + "_" + Date.now());
        var file = folder.createFile(blob);
        // [v62.1 최적화] 부모 폴더(GenieWorld_Archive)가 이미 Public View 권한을 상속하므로, 
        // 개별 파일 setSharing은 불필요하게 10초 이상 지연을 유발하므로 주석 처리하여 업로드 속도를 10배 이상 단축시킵니다!
        // file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoId = file.getId();
        console.log("[v44.199] 사진 저장 성공! ID:", photoId);
      } catch (err) {
        var errStr = err.toString();
        photoError = (errStr.indexOf("Permission") > -1) ? "드라이브 권한 승인 필요" : "저장 오류: " + errStr;
        console.error("Archive Drive Error:", errStr);
      }
    }
    
    sheet.appendRow([
      dateStr, 
      timeStr,
      payload.name,
      "'" + formattedPhone,
      payload.type,
      payload.item,
      payload.comment || "",
      photoId, 
      payload.score || 0
    ]);
    
    // [v45.3] 타입별 스탯 분류 자동화: 습관 -> 방어력(def), 퀘스트/식단 -> 수행력(perf)
    var statType = "perf"; 
    if (payload.type === "습관") statType = "def";

    recordActivityLog({
      phone: payload.phone, 
      name: payload.name, 
      type: payload.type, // "퀘스트" or "습관" or "식단"
      item: payload.item,
      action: "인증",
      score: payload.score || 0,
      statType: statType,
      photoId: photoId
    });

    // [v47.0] 아침 식단 Tier S (수분/차/단식) 인증 시 모닝 티 습관 연계 완료/인증 자동 지급!
    if (payload.type === "식단" && payload.item.indexOf("아침 식단") > -1 && payload.item.indexOf("Tier S") > -1) {
      // 1. 모닝 티 완료 (체크) - 2점 적립 (방어/회복력)
      recordActivityLog({
        phone: payload.phone, 
        name: payload.name, 
        type: "습관", 
        item: "모닝 티",
        action: "완료",
        score: 2,
        statType: "def"
      });
      // 2. 모닝 티 인증 (사진인증) - 3점 적립 (방어/회복력)
      recordActivityLog({
        phone: payload.phone, 
        name: payload.name, 
        type: "습관", 
        item: "모닝 티",
        action: "인증",
        score: 3,
        statType: "def",
        photoId: photoId
      });
    }

    // [v47.0] 저녁 식단 Tier S (저녁 단식 - 물/차만 섭취로 15점 획득) 시 나이트 컷 습관 완료 자동 지급!
    if (payload.type === "식단" && (payload.item.indexOf("저녁 식단") > -1 || payload.item.indexOf("저녁 단식") > -1) && payload.item.indexOf("Tier S") > -1 && payload.score === 15) {
      // 나이트 컷 완료 (체크) - 5점 적립 (방어/회복력)
      recordActivityLog({
        phone: payload.phone, 
        name: payload.name, 
        type: "습관", 
        item: "나이트 컷",
        action: "완료",
        score: 5,
        statType: "def"
      });
    }

    // [v46.10] 글리코겐 클리어 퀘스트 발동 조건 감시 (식단이고 Tier C인 경우)
    if (payload.type === "식단" && payload.item.indexOf("Tier C") > -1) {
      triggerGlycogenQuest(formattedPhone, payload.name);
    }

    var lastRow = sheet.getLastRow();

    return { 
      success: true, 
      photoId: photoId, 
      warning: photoError,
      receivedImageLen: (payload.image ? String(payload.image).length : 0),
      debugInfo: (photoId ? "SAVE_OK" : (photoError || "IMG_SKIP")) + " (To: " + ss.getName() + " > " + sheet.getName() + " [Row: " + lastRow + "])",
      serverVersion: "v44.203"
    };
  } catch (e) {
    return { success: false, error: "기록 저장 실패: " + e.toString() };
  }
}

/**
 * [대시보드] 특정 회원의 실시간 데이터(정보 + 점수) 가져오기 (v44.146)
 */
function getUserDashboardData(payload) {
  try {
    var rawPhone = String(payload.phone || "");
    var phone = normalizePhoneDigits(rawPhone); 
    if (!phone) return { error: "전화번호가 없습니다." };
    

    // [perf] 구글 초고속 캐시 서비스 확인 (0.05초)
    var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    
    // [v56.5] 릴레이 칭찬 섹션 영구 폐지로 관련 백엔드 실시간 알림/바톤 스캐너 로직 완전 제거!
    var praiseNotice = "";
    var praiseBatonSender = "";

    var cache = CacheService.getScriptCache();
    var cacheKey = "v62_dash_" + phone + "_" + todayStr;
    var cachedData = cache.get(cacheKey);
    if (cachedData) {
      try {
        var parsed = JSON.parse(cachedData);
        if (parsed && parsed.success) {
          Logger.log("⚡ 캐시 데이터 반환 완료 (0.01초): " + phone);
          parsed.cacheHit = true;
          parsed.praiseNotice = praiseNotice; // 캐시 반환 시에도 실시간 칭찬 알림 실시간 동기화!
          parsed.praiseBatonSender = praiseBatonSender; // [v55.0] 실시간 칭찬 바톤 연동!
          parsed.villageSettings = getVillageSettings();
          parsed.pillarNotice = getPillarNotice();
          return parsed;
        }
      } catch (e) {
        cache.remove(cacheKey);
      }
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 회원 정보 (등록현황 시트)
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    var memberInfo = { name: "모험가", tier: "씨앗", rank: "-" };
    if (regSheet) {
      var regData = regSheet.getDataRange().getDisplayValues();
      var regCols = getRegColumnIndices(regSheet);
      for (var i = 1; i < regData.length; i++) {
        var sheetPhone = normalizePhoneDigits(regData[i][regCols.phone]); 
        if (sheetPhone === phone || (phone.length >= 8 && sheetPhone.endsWith(phone.substring(phone.length - 8)))) {
          memberInfo.name = regData[i][regCols.name];
          memberInfo.tier = regData[i][regCols.membership] || "새싹";
          break;
        }
      }
    }

    // 회원명단 시트에서 목표체중 로드
    var targetWeight = 0;
    var memberSheet = ss.getSheetByName("회원명단");
    if (memberSheet) {
      var mData = memberSheet.getDataRange().getDisplayValues();
      var mCols = getMemberSheetColumnIndices(memberSheet);
      for (var i = 1; i < mData.length; i++) {
        var mPhone = normalizePhoneDigits(mData[i][mCols.phone]);
        if (mPhone === phone || (phone.length >= 8 && mPhone.endsWith(phone.substring(phone.length - 8)))) {
          targetWeight = parseWeightSafely(mData[i][mCols.targetWeight]);
          break;
        }
      }
    }

    // [v45.8] 일일_활동_기록 (10컬럼 집약 체계)
    var summarySheet = ss.getSheetByName("일일_활동_기록") || ss.insertSheet("일일_활동_기록");
    var doneList = []; 
    var stats = { 
      weekly: { health: 0, perf: 0, def: 0 },
      monthly: { health: 0, perf: 0, def: 0 }
    };
    var activityLifetime = 0;
    var activitySeason = 0;
    var activityWeekly = 0;
    var achievedBonuses = [];
    
    var now = new Date();
    
    // [v46.43] 주간: 매주 목요일 00:00:00 기점 (수요일 자정 마감)
    var day = now.getDay(); // 0(일)~6(토)
    var diffToThu = (day + 3) % 7; 
    var startOfWeek = new Date(now.getFullYear(), now.getMonth(), now.getDate() - diffToThu);
    startOfWeek.setHours(0, 0, 0, 0);

    // [v45.9] 월간(시즌): 매월 4째주 수요일 기점 (33챌린지 사이클)
    function get4thWednesday(year, month) {
      var firstDay = new Date(year, month, 1);
      var firstWed = new Date(year, month, 1 + ((3 - firstDay.getDay() + 7) % 7));
      return new Date(year, month, firstWed.getDate() + 21);
    }
    
    var currentMonth4thWed = get4thWednesday(now.getFullYear(), now.getMonth());
    currentMonth4thWed.setHours(0, 0, 0, 0);
    
    var startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0, 0);
    
    var startOfSeason;
    if (now >= currentMonth4thWed) {
      startOfSeason = currentMonth4thWed;
    } else {
      startOfSeason = get4thWednesday(now.getFullYear(), now.getMonth() - 1);
      startOfSeason.setHours(0, 0, 0, 0);
    }

    var data = summarySheet.getDataRange().getDisplayValues();
    for (var j = 1; j < data.length; j++) {
      var recPhone = normalizePhoneDigits(data[j][1]);
      if (recPhone === phone) {
        var recDateStr = String(data[j][0]);
        var dMatch = recDateStr.match(/(\d{4})[^\d]+(\d{1,2})[^\d]+(\d{1,2})/);
        if (dMatch) {
          var recDate = new Date(dMatch[1], parseInt(dMatch[2], 10) - 1, parseInt(dMatch[3], 10));
          var rowTotal = Number(data[j][9] || 0); // J열(10번째): 총점
          activityLifetime += rowTotal;
          
          // 월간(시즌) 합산: 1일부터 말일까지로 행동 스탯 누적 교정
          if (recDate >= startOfMonth) {
            activitySeason += rowTotal;
            stats.monthly.perf += (Number(data[j][5]) || 0);
            stats.monthly.def += (Number(data[j][6]) || 0);
            stats.monthly.health += (Number(data[j][3]) || 0) + (Number(data[j][4]) || 0) + (Number(data[j][7]) || 0);
          }
          // 주간 합산
          if (recDate >= startOfWeek) {
            activityWeekly += rowTotal;
            stats.weekly.perf += (Number(data[j][5]) || 0);
            stats.weekly.def += (Number(data[j][6]) || 0);
            stats.weekly.health += (Number(data[j][3]) || 0) + (Number(data[j][4]) || 0) + (Number(data[j][7]) || 0);

            // [v46.4] 이번 주 획득한 체력보너스 탐색
            var rowDetails = String(data[j][8] || "");
            if (rowDetails.indexOf("주간 3회 출석 보너스") > -1) achievedBonuses.push(3);
            if (rowDetails.indexOf("주간 4회 출석 보너스") > -1) achievedBonuses.push(4);
            if (rowDetails.indexOf("주간 5회 출석 보너스") > -1) achievedBonuses.push(5);
            if (rowDetails.indexOf("주간 6회 출석 보너스") > -1) achievedBonuses.push(6);

            if (Utilities.formatDate(recDate, "GMT+9", "yyyy-MM-dd") === todayStr) {
              var details = String(data[j][8] || "");
              var items = details.split(", ");
              items.forEach(function(item) { 
                var cleanItem = item.split("(")[0].replace(/\[.*?\]\s*/, "").trim();
                if (cleanItem) doneList.push(cleanItem); 
              });
            }
          }
        }
      }
    }

    // [v47.0] 실시간 로그인 감지 및 자동 점수 지급 엔진 (로그아웃 불필요!)
    var hasLoginToday = false;
    doneList.forEach(function(item) {
      if (item.indexOf("로그인 체크") > -1) {
        hasLoginToday = true;
      }
    });

    var isFirstLoginToday = false;
    if (!hasLoginToday && memberInfo.name !== "모험가") {
      isFirstLoginToday = true;
      // 오늘 첫 로그인 점수 자동 기록
      recordActivityLog({
        phone: phone,
        name: memberInfo.name,
        type: "로그인",
        item: "로그인 체크",
        score: 5,
        statType: "def"
      });
      // 실시간 데이터 즉시 업데이트
      activityLifetime += 5;
      activityWeekly += 5;
      activitySeason += 5;
      stats.weekly.def += 5;
      stats.monthly.def += 5;
      doneList.push("로그인 체크");
    }

    // [v46.37] 주간 랭킹 발표 및 잠금 기준선 (목요일 12:00) & 측정 마감선 (수요일 24:00) 연산 엔진
    var currentThursdayNoon = new Date(now.getTime());
    var dayOfWeek = currentThursdayNoon.getDay(); // 0(일)~6(토)
    var diffToThu = 4 - dayOfWeek;
    currentThursdayNoon.setDate(currentThursdayNoon.getDate() + diffToThu);
    currentThursdayNoon.setHours(12, 0, 0, 0);
    
    // 만약 현재가 이번 주 목요일 12:00 이전이라면, 지난주 목요일 12:00을 기준선으로 설정하여 지난주 순위를 고정 표출합니다.
    var targetThursdayNoon = new Date(currentThursdayNoon.getTime());
    if (now < currentThursdayNoon) {
      targetThursdayNoon.setDate(targetThursdayNoon.getDate() - 7);
    }
    
    // 측정일 마감 기준선: 수요일 24:00 (즉 목요일 00:00:00)
    var measurementDeadline = new Date(targetThursdayNoon.getTime());
    measurementDeadline.setDate(measurementDeadline.getDate() - 1); // 목요일 -> 수요일
    measurementDeadline.setHours(24, 0, 0, 0); // 수요일 24:00 (목요일 00:00:00)
    
    // 제출(업로드) 마감 기준선: 목요일 12:00
    var submissionDeadline = targetThursdayNoon;
    var lastWed = measurementDeadline; // 하위 호환성 유지용 선언

    // [v65.70] 당월 1일 ~ 수학적 1주차 마감 수요일 Milestone 계산
    var tYear = now.getFullYear();
    var tMonth = now.getMonth();
    var startOfMonth = new Date(tYear, tMonth, 1, 0, 0, 0, 0);
    var firstDayOfWeek = startOfMonth.getDay(); // 0(일)~6(토)
    var diffToWed = (3 - firstDayOfWeek + 7) % 7;
    var baselineDeadline = new Date(tYear, tMonth, 1 + diffToWed);
    if (typeof isCenterHoliday === "function" && isCenterHoliday(baselineDeadline)) {
      baselineDeadline.setDate(baselineDeadline.getDate() + 1); // 공휴일 시 목요일 연장
    }
    baselineDeadline.setHours(23, 59, 59, 999);

    var firstEverInbody = null;
    var activeInbodySnapshot = null;
    var season0WeekInbody = null;
    var latestInbody = null; // [v46.41] 실시간 개인용 최신 인바디 스냅샷

    var inbodySheet = ss.getSheetByName("33챌린지_인바디");
    if (inbodySheet) {
      var inData = inbodySheet.getDataRange().getValues();
      for (var k = 1; k < inData.length; k++) {
        var recPhone = normalizePhoneDigits(inData[k][2]);
        if (recPhone === phone) {
          var iDate = new Date(inData[k][0]);
          if (isNaN(iDate.getTime())) continue;
          
          var remarksStr = String(inData[k][7] || "");
          var record = {
            date: iDate,
            weight: isNaN(Number(inData[k][3])) ? 0 : Number(inData[k][3]),
            muscle: isNaN(Number(inData[k][4])) ? 0 : Number(inData[k][4]),
            fat: isNaN(Number(inData[k][5])) ? 0 : Number(inData[k][5]),
            remarks: remarksStr
          };
          
          // [v46.38] 전용 9번째 열(I열)에서 등록일(Record Date) 추출, 없으면 측정일(A열)과 동일하게 매칭!
          var uploadDate = iDate;
          if (inData[k].length > 8 && inData[k][8]) {
            var rawUpload = inData[k][8];
            var uObj = (rawUpload instanceof Date) ? rawUpload : new Date(rawUpload);
            if (!isNaN(uObj.getTime())) {
              uploadDate = uObj;
            }
          }
          
          if (!firstEverInbody || iDate < firstEverInbody.date) {
            firstEverInbody = record;
          }
          
          // 1. 측정일은 수요일 24:00 (목요일 00:00) 이전
          // 2. 제출일은 목요일 12:00 (정오) 이전인 기록만 이번 랭킹 스냅샷에 포함!
          if (iDate <= measurementDeadline && uploadDate <= submissionDeadline) {
            if (!activeInbodySnapshot || iDate > activeInbodySnapshot.date) {
              activeInbodySnapshot = record;
            }
          }
          // [v65.70] 당월 1일 ~ 수학적 1주차 마감 수요일 사이에 등록된 최초 기록을 월초 기준 인바디로 설정!
          if (iDate >= startOfSeason && iDate <= baselineDeadline) {
            if (!season0WeekInbody || iDate < season0WeekInbody.date) {
              season0WeekInbody = record;
            }
          }
          // [v46.41] 실시간 갱신을 위해 랭킹 마감 여부 상관없이 무조건 최신 기록으로 할당!
          if (!latestInbody || iDate > latestInbody.date) {
            latestInbody = record;
          }
        }
      }
    }

    // 인바디 개선 점수 연산 헬퍼 호출
    function calculateInbodyScoreHelper(first, current, targetWeight, scoreType) {
      if (!first || !current) return 0;
      
      var fW = Number(first.weight) || 0;
      var cW = Number(current.weight) || 0;
      var fM = Number(first.muscle) || 0;
      var cM = Number(current.muscle) || 0;
      var fF = Number(first.fat) || 0;
      var cF = Number(current.fat) || 0;
      
      var score = 0;
      var diffW = Number((fW - cW).toFixed(2)) || 0;
      var diffM = Number((cM - fM).toFixed(2)) || 0;
      var diffFat = Number((fF - cF).toFixed(1)) || 0;
      
      if (diffW > 0) score += (diffW * 10) * 50;
      if (diffM > 0) score += (diffM * 10) * 200;
      if (diffFat > 0) score += (diffFat * 10) * 100;
      
      var firstFatMass = fW * (fF / 100);
      var currentFatMass = cW * (cF / 100);
      var fatLossRate = 0;
      if (firstFatMass > 0 && !isNaN(firstFatMass)) {
        fatLossRate = ((firstFatMass - currentFatMass) / firstFatMass) * 100;
      }
      
      if (diffW >= 8.0 || (fatLossRate >= 20.0 && !isNaN(fatLossRate))) {
        score += 1500;
      }
      
      // 🏆 체성분 명품 유지 보너스 판정 엔진 (±0.5kg 정교화)
      if (targetWeight && targetWeight > 0) {
        if (scoreType === "monthly") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        } else if (scoreType === "lifetime") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        }
      }
      
      return isNaN(score) ? 0 : score;
    }

    // 평생 최초와 최신 기록이 동일한 기록 1개뿐인 경우(자기 대조) 점수 0점
    var inbodyLifetimeScore = 0;
    if (firstEverInbody && latestInbody && firstEverInbody !== latestInbody) {
      inbodyLifetimeScore = calculateInbodyScoreHelper(firstEverInbody, latestInbody, targetWeight, 'lifetime');
    }
    
    // [v46.42] 주간 및 시즌(월간) 인바디 점수 정밀 계산 (v65.70 strict rules)
    var inbodyWeeklyScore = 0;
    var inbodySeasonScore = 0;
    
    var currentInbody = latestInbody || activeInbodySnapshot;
    if (currentInbody) {
      // (1) 주간 인바디 점수: 이번 주(startOfWeek 이후)에 측정된 기록이 있을 때만 반영!
      if (currentInbody.date >= startOfWeek) {
        // 이번 주 목요일 이전의 가장 최근 기록이되, 목요일 기준 7일(7.5일 마진) 이내인 지난주 기록만 탐색!
        var prevBeforeThisWeek = null;
        if (inbodySheet && inData) {
          for (var k = 1; k < inData.length; k++) {
            var recPhone = normalizePhoneDigits(inData[k][2]);
            if (recPhone === phone) {
              var iDate = new Date(inData[k][0]);
              if (iDate < startOfWeek) {
                var timeDiff = startOfWeek.getTime() - iDate.getTime();
                var diffDays = timeDiff / (1000 * 60 * 60 * 24);
                if (diffDays <= 7.5) { // 7.5일 이내 지난주 기록 한정
                  if (!prevBeforeThisWeek || iDate > prevBeforeThisWeek.date) {
                    prevBeforeThisWeek = {
                      date: iDate,
                      weight: Number(inData[k][3] || 0),
                      muscle: Number(inData[k][4] || 0),
                      fat: Number(inData[k][5] || 0)
                    };
                  }
                }
              }
            }
          }
        }
        // 지난주 기록(prevBeforeThisWeek)이 있을 때만 주간 점수 부여! 없으면 무조건 0점.
        if (prevBeforeThisWeek) {
          inbodyWeeklyScore = calculateInbodyScoreHelper(prevBeforeThisWeek, currentInbody, targetWeight, 'weekly');
        }
      }
      
      // (2) 시즌(월간) 인바디 점수: 월초 첫 주 기준 기록(season0WeekInbody)이 있고, 본인 최신 기록이 다를 때만 반영! (엄격 마감 규칙 적용)
      if (season0WeekInbody && currentInbody && season0WeekInbody !== currentInbody) {
        if (isDateInLastWeekMonToWed(currentInbody.date)) {
          inbodySeasonScore = calculateInbodyScoreHelper(season0WeekInbody, currentInbody, targetWeight, 'monthly');
        } else {
          inbodySeasonScore = 0;
        }
      }
    }


    // [v46.35] 연속 미출석(결석) 페널티 계산 (출석 시 즉시 부활!)
    var lastAttendanceDate = null;
    var logSheet = ss.getSheetByName("출석기록");
    if (logSheet) {
      var logData = logSheet.getDataRange().getValues();
      var logCols = getAttendanceColumnIndices(logSheet);
      for (var l = logData.length - 1; l >= 1; l--) {
        var sheetPhone = normalizePhoneDigits(logData[l][logCols.phone]);
        if (sheetPhone === phone) {
          var dRaw = logData[l][logCols.date];
          if (dRaw) {
            var dObj = (dRaw instanceof Date) ? dRaw : new Date(dRaw);
            if (!isNaN(dObj.getTime())) {
              lastAttendanceDate = dObj;
              break;
            }
          }
        }
      }
    }
    
    var inactiveDays = 0;
    var inactivityPenalty = 0;
    if (lastAttendanceDate) {
      var midnightNow = new Date();
      midnightNow.setHours(0,0,0,0);
      var midnightLast = new Date(lastAttendanceDate.getTime());
      midnightLast.setHours(0,0,0,0);
      
      inactiveDays = calculateInactiveDays(midnightLast, midnightNow);
      
      if (inactiveDays > 3) {
        inactivityPenalty = (inactiveDays - 3) * 100;
        if (inactivityPenalty > 3000) inactivityPenalty = 3000;
      }
    }

    // 최종 스코어 보정 및 디버프 차감
    var scores = {};
    scores.lifetime = Math.max(0, activityLifetime + inbodyLifetimeScore - inactivityPenalty);
    scores.weekly = Math.max(0, activityWeekly + inbodyWeeklyScore - inactivityPenalty);
    scores.season = Math.max(0, activitySeason + inbodySeasonScore - inactivityPenalty);

    stats.weekly.health += inbodyWeeklyScore;
    stats.monthly.health += inbodySeasonScore;

    // 4. 티어 및 랭킹 (기존 로직 유지)
    var tiers = [
      { name: "씨앗 🌱", min: 0, max: 1000 },
      { name: "새싹 🌿", min: 1001, max: 3000 },
      { name: "나무 🌳", min: 3001, max: 8000 },
      { name: "꽃 🌸", min: 8001, max: 15000 },
      { name: "꿈나무 요정 🧚‍♂️", min: 15001, max: 30000 },
      { name: "전설의 점퍼 👑", min: 30001, max: 60000 },
      { name: "지니 월드 수호신 🌌", min: 60001, max: 9999999 }
    ];
    var currentTierIndex = 0;
    var currentTier = tiers[0];
    for (var t = 0; t < tiers.length; t++) {
      if (scores.lifetime >= tiers[t].min) {
        currentTier = tiers[t];
        currentTierIndex = t;
      } else {
        break;
      }
    }
    
    var nextTier = "지니 월드 수호신 🌌";
    var evolution = 100;
    if (currentTierIndex < tiers.length - 1) {
      var next = tiers[currentTierIndex + 1];
      nextTier = next.name;
      var tierRange = next.min - currentTier.min;
      var earnedInTier = scores.lifetime - currentTier.min;
      evolution = Math.min(Math.floor((earnedInTier / tierRange) * 100), 100);
      if (evolution < 0) evolution = 0;
    }

    var questStatus = getActiveQuestStatus(phone, ss, logData, memberInfo.name);

    // 5. [v47.1] 실시간 전체 회원 웰니스 랭킹 산출 엔진 (명예의 전당의 인바디 포함 온전 누적합과 100% 동기화!)
    var rank = "-";
    try {
      var hof = getHallOfFameData();
      if (hof && hof.success && hof.data && hof.data.total) {
        var cleanName = memberInfo.name.replace(/\d{4}$/, "").trim();
        var myRankEntry = hof.data.total.find(function(entry) {
          var entryName = String(entry.name || "").trim();
          return entryName === cleanName;
        });
        if (myRankEntry) {
          rank = myRankEntry.rank;
        }
      }
    } catch (err) {
      Logger.log("실시간 랭킹 연산 오류: " + err.toString());
    }

    // [v59.0] 오늘 첫 로그인 시 알림 쪽지 생성 트리거
    if (isFirstLoginToday) {
      var rawName = String(memberInfo.name || "회원").replace(/\d{4}$/, "").trim();
      var friendlyName = getFriendlyName(memberInfo.name);
      
      if (inactivityPenalty > 0) {
        // 디버프 경고 쪽지는 정확한 전달과 격식을 위해 성을 포함한 전체 이름 사용
        sendPersonalNotification(
          phone,
          "debuff",
          "연속 미출석 에너지 방전 디버프 발생!",
          rawName + " 회원님! 클럽 출석을 하지 않으신 지 연속 " + inactiveDays + "일이 경과하여 웰니스 에너지가 방전되었습니다. 누적 점수에서 -" + inactivityPenalty + " EXP가 차감되었습니다. 💡 오늘 클럽에 출석하시면 즉시 100% 원상 복원됩니다!"
        );
      } else {
        // 일상 웰컴 쪽지는 친근함을 극대화하기 위해 성을 뗀 이름(예: 미진님) 사용
        sendPersonalNotification(
          phone,
          "welcome",
          "오늘 첫 로그인! 웰니스 보너스 지급 완료",
          friendlyName + "님! 오늘 하루도 힘차게 시작해봐요! 로그인 보너스로 +5 EXP(회복력)가 즉시 지급되었습니다. ⚔️"
        );
      }
    }

    var personalNotiResult = getPersonalNotifications({ phone: phone });

    var result = {
      success: true,
      doneList: doneList,
      name: memberInfo.name,
      tier: currentTier.name,
      nextTier: nextTier,
      evolution: evolution,
      totalScore: scores.lifetime, 
      seasonScore: scores.season,
      weeklyScore: scores.weekly,
      stats: stats, // { weekly, monthly } 데이터
      rank: rank,
      weeklyTargets: { health: 1500, perf: 1000, def: 500 },
      monthlyTargets: { health: 6000, perf: 4000, def: 2000 },
      quests: questStatus,
      achievedBonuses: achievedBonuses,
      inactiveDays: inactiveDays,
      inactivityPenalty: inactivityPenalty,
      isFirstLoginToday: isFirstLoginToday,
      praiseNotice: praiseNotice, // 실시간 칭찬 알림 전달!
      praiseBatonSender: praiseBatonSender, // [v55.0] 실시간 칭찬 바톤 연동!
      villageSettings: getVillageSettings(),
      pillarNotice: getPillarNotice(),
      notifications: personalNotiResult.success ? personalNotiResult.notifications : []
    };
    
    result.cacheHit = false;
    result.cacheError = null;
    
    // [perf] 구글 캐시에 3분(180초) 동안 저장
    try {
      cache.put(cacheKey, JSON.stringify(result), 180);
    } catch (err) {
      Logger.log("캐시 저장 오류: " + err.toString());
      result.cacheError = err.toString();
    }
    
    return result;
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * [perf] 회원의 대시보드 캐시를 즉시 파괴하는 유틸리티 (실시간 데이터 갱신용)
 */
function clearUserDashboardCache(phone) {
  if (!phone) return;
  var cleanPhone = String(phone).replace(/[^0-9]/g, "");
  try {
    var cache = CacheService.getScriptCache();
    var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    cache.remove("v62_dash_" + cleanPhone + "_" + todayStr);
    Logger.log("⚡ 대시보드 캐시 강제 삭제 완료: " + cleanPhone);
  } catch (e) {
    Logger.log("캐시 삭제 오류: " + e.toString());
  }
}

/**
 * [v44.160] 인바디 기록 저장 및 v39 자동 점수 계산 (주간/월간/토탈 3단 점수 정교화 패치 v65.70)
 */
function submitInBodyRecord(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지_인바디") || ss.insertSheet("33챌린지_인바디");
    
    // [마이그레이션 가드] 기존 9열 포맷("변화점수")을 11열 포맷("주간/월간/토탈변화점수")으로 자동 변환
    if (sheet.getLastRow() > 0) {
      var firstRowValues = sheet.getRange(1, 1, 1, 9).getValues()[0];
      var headerG = String(firstRowValues[6] || "").trim();
      if (headerG === "변화점수" || headerG === "변화 점수") {
        sheet.getRange(1, 1, 1, 11).setValues([[
          "측정일", "회원명", "연락처", "체중", "골격근량", "체지방률", 
          "주간변화점수", "월간변화점수", "토탈변화점수", "비고", "등록일"
        ]]);
        SpreadsheetApp.flush();
      }
    } else {
      // 신규 시트인 경우 헤더 생성
      sheet.appendRow([
        "측정일", "회원명", "연락처", "체중", "골격근량", "체지방률", 
        "주간변화점수", "월간변화점수", "토탈변화점수", "비고", "등록일"
      ]);
    }
    
    var name = payload.name;
    var phone = formatPhoneNumber(payload.phone).replace(/[^0-9]/g, "");
    var weight = Number(payload.weight);
    var muscle = Number(payload.muscle);
    var fat = Number(payload.fat);

    // 회원명단 시트에서 목표체중 로드
    var targetWeight = 0;
    var memberSheet = ss.getSheetByName("회원명단");
    if (memberSheet) {
      var mData = memberSheet.getDataRange().getDisplayValues();
      var mCols = getMemberSheetColumnIndices(memberSheet);
      for (var i = 1; i < mData.length; i++) {
        var mPhone = normalizePhoneDigits(mData[i][mCols.phone]);
        if (mPhone === phone || (phone.length >= 8 && mPhone.endsWith(phone.substring(phone.length - 8)))) {
          targetWeight = parseWeightSafely(mData[i][mCols.targetWeight]);
          break;
        }
      }
    }
    
    var dateValue = new Date();
    if (payload.customDate) {
      var parts = payload.customDate.split('-');
      if (parts.length === 3) {
        var now = new Date();
        dateValue = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), now.getHours(), now.getMinutes(), now.getSeconds());
      }
    }
    
    var targetDateStr = Utilities.formatDate(dateValue, "GMT+9", "yyyy-MM-dd");

    // 1. 전체 이력 조회 및 동일 날짜 중복 검사
    var data = sheet.getDataRange().getValues();
    var existingRowIndex = -1;
    var userRecords = [];
    
    // G열이 주간변화점수 등 11열 포맷으로 매핑되었으므로, 헤더가 아닌 데이터 행을 안전하게 정밀 스캔
    for (var i = 1; i < data.length; i++) {
      var rowPhone = formatPhoneNumber(data[i][2]).replace(/[^0-9]/g, "");
      if (rowPhone === phone) {
        var rowDate = new Date(data[i][0]);
        if (isNaN(rowDate.getTime())) continue;
        
        var rowDateStr = Utilities.formatDate(rowDate, "GMT+9", "yyyy-MM-dd");
        if (rowDateStr === targetDateStr) {
          existingRowIndex = i + 1; // 동일 날짜 중복
        }
        
        userRecords.push({
          date: rowDate,
          weight: Number(data[i][3] || 0),
          muscle: Number(data[i][4] || 0),
          fat: Number(data[i][5] || 0)
        });
      }
    }
    
    if (existingRowIndex > -1) {
      return { 
        success: false, 
        error: "선택하신 측정 날짜(" + targetDateStr + ")에 이미 등록된 인바디 기록이 존재합니다!\n\n중복 등록은 불가하오니 날짜를 다시 확인해 주세요. 만약 수치 오입력으로 기존 기록의 수정을 원하시는 경우 웰니스 코치에게 직접 요청해 주시기 바랍니다! 😊" 
      };
    }
    
    // 2. 신규 등록 데이터를 포함한 가상의 이력 시뮬레이션용 배열 구성 (날짜순 정렬)
    var currentInbody = { date: dateValue, weight: weight, muscle: muscle, fat: fat };
    var simulatedRecords = userRecords.slice();
    simulatedRecords.push(currentInbody);
    simulatedRecords.sort(function(a, b) { return a.date - b.date; });
    
    // 3. 인바디 점수 연산 공통 헬퍼
    function calculateInbodyScoreHelper(first, current, targetWeight, scoreType) {
      if (!first || !current) return 0;
      var score = 0;
      var fW = Number(first.weight) || 0;
      var cW = Number(current.weight) || 0;
      var fM = Number(first.muscle) || 0;
      var cM = Number(current.muscle) || 0;
      var fF = Number(first.fat) || 0;
      var cF = Number(current.fat) || 0;
      
      var diffW = Number((fW - cW).toFixed(2)) || 0;
      var diffM = Number((cM - fM).toFixed(2)) || 0;
      var diffFat = Number((fF - cF).toFixed(1)) || 0;
      
      if (diffW > 0) score += (diffW * 10) * 50;
      if (diffM > 0) score += (diffM * 10) * 200;
      if (diffFat > 0) score += (diffFat * 10) * 100;
      
      var firstFatMass = fW * (fF / 100);
      var currentFatMass = cW * (cF / 100);
      var fatLossRate = 0;
      if (firstFatMass > 0 && !isNaN(firstFatMass)) {
        fatLossRate = ((firstFatMass - currentFatMass) / firstFatMass) * 100;
      }
      if (diffW >= 8.0 || (fatLossRate >= 20.0 && !isNaN(fatLossRate))) {
        score += 1500;
      }
      
      // 🏆 체성분 명품 유지 보너스 판정 엔진 (±0.5kg 정교화)
      if (targetWeight && targetWeight > 0) {
        if (scoreType === "monthly") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        } else if (scoreType === "lifetime") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        }
      }
      
      return isNaN(score) ? 0 : score;
    }
    
    // 4. 날짜 기준점 계산
    // (1) 이번 주 목요일 구하기
    var day = dateValue.getDay();
    var diffToThu = (day + 3) % 7; 
    var startOfWeek = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate() - diffToThu);
    startOfWeek.setHours(0, 0, 0, 0);
    
    // (2) 월초 시작일 및 수학적 1주차 수요일 마감 마일스톤 구하기
    var year = dateValue.getFullYear();
    var month = dateValue.getMonth();
    var startOfMonth = new Date(year, month, 1, 0, 0, 0, 0);
    
    // [원장님 선포 규칙] 1일 이후 도래하는 첫 수요일
    var firstDayOfWeek = startOfMonth.getDay(); // 0(일)~6(토)
    var diffToWed = (3 - firstDayOfWeek + 7) % 7;
    var baselineDeadline = new Date(year, month, 1 + diffToWed);
    var hasWedHoliday = isCenterHoliday(baselineDeadline);
    if (hasWedHoliday) {
      baselineDeadline.setDate(baselineDeadline.getDate() + 1); // 공휴일 시 목요일 연장
    }
    baselineDeadline.setHours(23, 59, 59, 999);
    
    // 5. 기준 대조 기록 필터링
    var firstEver = null;
    var prevBeforeThisWeek = null;
    var baselineRecord = null;
    
    for (var rIdx = 0; rIdx < simulatedRecords.length; rIdx++) {
      var r = simulatedRecords[rIdx];
      
      // A. 평생 최초 기록 구하기
      if (!firstEver || r.date < firstEver.date) {
        firstEver = r;
      }
      
      // B. 지난주 인바디 구하기: 이번 주 목요일 이전의 가장 최근 기록이되, 목요일 기준 7일(7.5일 마진) 이내인 것!
      if (r.date < startOfWeek) {
        var timeDiff = startOfWeek.getTime() - r.date.getTime();
        var diffDays = timeDiff / (1000 * 60 * 60 * 24);
        if (diffDays <= 7.5) { // 지난주 측정일 범위 내
          if (!prevBeforeThisWeek || r.date > prevBeforeThisWeek.date) {
            prevBeforeThisWeek = r;
          }
        }
      }
      
      // C. 월초 기준 인바디 구하기: 당월 1일 ~ 1주차 마감 수요일 사이의 첫 기록
      if (r.date >= startOfMonth && r.date <= baselineDeadline) {
        if (!baselineRecord || r.date < baselineRecord.date) {
          baselineRecord = r;
        }
      }
    }
    
    // 6. 3단 점수 정밀 연산
    var weeklyScore = 0;
    var monthlyScore = 0;
    var totalScore = 0;
    
    // 최초 기록과 본인의 기록이 단 1개뿐인 경우(자기 대조) 점수 0점
    if (firstEver && firstEver !== currentInbody) {
      totalScore = calculateInbodyScoreHelper(firstEver, currentInbody, targetWeight, 'lifetime');
    }
    
    if (prevBeforeThisWeek) {
      weeklyScore = calculateInbodyScoreHelper(prevBeforeThisWeek, currentInbody, targetWeight, 'weekly');
    }
    
    if (baselineRecord && baselineRecord !== currentInbody) {
      if (isDateInLastWeekMonToWed(currentInbody.date)) {
        monthlyScore = calculateInbodyScoreHelper(baselineRecord, currentInbody, targetWeight, 'monthly');
      } else {
        monthlyScore = 0;
      }
    }
    
    // 7. 시트에 3단 확장 포맷으로 영구 기입
    var uploadDate = new Date();
    
    var finalRemarks = payload.remarks || "";
    if (targetWeight && targetWeight > 0) {
      var cW = Number(currentInbody.weight) || 0;
      if (cW <= targetWeight + 0.5) {
        var bonusMsg = "[🏆 체성분 명품 유지 보너스 적용 (목표: " + targetWeight + "kg / 현재: " + cW.toFixed(1) + "kg)]";
        if (finalRemarks) {
          finalRemarks += " " + bonusMsg;
        } else {
          finalRemarks = bonusMsg;
        }
      }
    }

    sheet.appendRow([
      dateValue,              // Column A: 측정일 (Measurement Date)
      name,                   // Column B: 회원명
      "'" + phone,            // Column C: 연락처
      weight,                 // Column D: 체중
      muscle,                 // Column E: 골격근량
      fat,                    // Column F: 체지방률
      weeklyScore,            // Column G: 주간변화점수 [NEW]
      monthlyScore,           // Column H: 월간변화점수 [NEW]
      totalScore,             // Column I: 토탈변화점수 [NEW]
      finalRemarks,           // Column J: 비고
      uploadDate              // Column K: 등록일 (Record Date)
    ]);
    
    // [perf] 인바디 업로드 완료 시 대시보드 캐시 즉시 파괴
    clearUserDashboardCache(phone);
    
    return { 
      success: true, 
      weeklyScore: weeklyScore, 
      monthlyScore: monthlyScore, 
      totalScore: totalScore,
      targetWeight: targetWeight,
      weight: weight
    };
  } catch (e) { 
    return { success: false, error: e.toString() }; 
  }
}


/**
 * [입장] 전화번호 뒷자리로 회원 검색 (키오스크 스타일 v44.148)
 */
// [이동 완료] searchMembersByDigits 함수는 Bridge_Login.gs로 이동되었습니다.


// [이동 완료] doGet, doPost, handleRequest 및 시트 컬럼 인덱스 매퍼 함수들은 Common_Utils.gs로 이동되었습니다.

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
          regSheet.getRange(i + 1, 11).setValue(expireInfo); // [이동 완료] 관련 함수들은 각각의 파일로 이동됨 
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
    workLogSheet.appendRow(["날짜", "작성자", "점핑목록", "근력목록", "09시", "10시", "17시", "18시", "19시", "20시", "총출석", "점핑인원", "테라피인원", "특이사항", "시설파손"]);
    workLogSheet.getRange("A1:O1").setFontWeight("bold").setBackground("#fff3cd");
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
// 6. 테라피 예약 시스템 백엔드 API (이동 완료)
// ──────────────────────────────────────────────
// [이동 완료] getAnnouncements() 함수는 Bridge_Login.gs 파일로 이동되었습니다.
// [이동 완료] getTodaySummary(), getTodayTimetable(), getRoomStatus(), getMyReservations(), getMemberIDList(), submitReservation() 함수들은 Reservation.gs 파일로 이동되었습니다.

// ──────────────────────────────────────────────
// 7. 관리자 모바일 대시보드 백엔드 API (admin.html 연동)
// ──────────────────────────────────────────────

// 백엔드 메인 안전 제어 밸브 (true: 초고속 신형 로직 / false: 100% 검증 레거시 로직)
var USE_FAST_DASHBOARD = true;

/**
// [이동 완료] formatPhoneSafely, formatTimeSafely, formatDateSafely, parseDateTimeSafely 함수들은 Common_Utils.gs로 이동되었습니다.

/**
 * 🏰 관리자 대시보드 데이터를 불러오기 위한 메인 라우터 (Dual-Core Switch)
 */
function getAdminDashboardData() {
  try {
    if (USE_FAST_DASHBOARD === false) {
      return getAdminDashboardData_Legacy();
    }
    return getAdminDashboardData_Fast();
  } catch (e) {
    return { error: "서버 메인 제어 오류: " + e.toString() };
  }
}

/**
 * ⚡ [최적화 버전] 3~5배 더 빠른 초고속 대시보드 리더
 */
function getAdminDashboardData_Fast() {
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
    
    // 1. 출석 로그 조회 (getValues()로 초고속 조회 및 3중 필터 보정)
    var logData = logSheet.getDataRange().getValues();
    var cols = getAttendanceColumnIndices(logSheet);
    
    for (var i = 1; i < logData.length; i++) {
      var dateRaw = logData[i][cols.date]; // 날짜 열 (객체 또는 문자열)
      if (!dateRaw) continue;
      
      var logDateStr = formatDateSafely(dateRaw);
      
      if (logDateStr === todayStr) {
        var memberName = logData[i][cols.name]; 
        var memberPhone = formatPhoneSafely(logData[i][cols.phone]);
        var membership = logData[i][cols.type]; 
        var status = String(logData[i][cols.status] || "").trim(); 
        var reason = String(logData[i][cols.reason] || "");
        
        // 시간 오차 및 형식 해결: 32분 저주/오타 강제 정화
        var rawTimeVal = logData[i][cols.inTime];
        var timeOnly = formatTimeSafely(rawTimeVal);
        
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
    
    // 2. 테라피 예약 매칭 및 예열 알림 체크 (getValues()로 초고속 조회)
    var resData = resSheet.getDataRange().getValues();
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
      var rDateVal = resData[k][resCols.date];
      if (!rDateVal) continue;
      
      var rDateNum = "";
      if (rDateVal instanceof Date) {
        rDateNum = Utilities.formatDate(rDateVal, "GMT+9", "yyyyMMdd");
      } else {
        var rDateRaw = String(rDateVal).trim();
        rDateNum = rDateRaw.replace(/[^0-9]/g, "");
        if (rDateNum.length > 8) rDateNum = rDateNum.substring(0, 8);
        if (rDateNum.length < 8) {
           var dParts = rDateRaw.split(/[^0-9]/).filter(Boolean);
           if (dParts.length >= 3) {
              rDateNum = dParts[0] + (dParts[1].length === 1 ? "0"+dParts[1] : dParts[1]) + (dParts[2].length === 1 ? "0"+dParts[2] : dParts[2]);
           }
        }
      }
      
      if (rDateNum === todayNum) {
        var rName = String(resData[k][resCols.name] || "").trim();
        var rTimeVal = resData[k][resCols.time];
        if (!rName || !rTimeVal) continue;
        
        var rTimeShort = formatTimeSafely(rTimeVal);
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
        var resHour = parseInt(rParts[0], 10);
        var resMinVal = parseInt(rParts[1], 10);
        var resTotalMin = (resHour * 60) + resMinVal;
 
        var nowKSTStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
        var nowParts = nowKSTStr.split(":");
        var nowTotalMin = (parseInt(nowParts[0], 10) * 60) + parseInt(nowParts[1], 10);
 
        for (var a=0; a<activeJumping.length; a++) {
          var mPhone = formatPhoneSafely(activeJumping[a].phone);
          var rPhone = formatPhoneSafely(resData[k][1]);
          var isPhoneMatch = (rPhone === mPhone) || (rPhone.length >= 8 && mPhone.length >= 8 && rPhone.slice(-8) === mPhone.slice(-8));
 
          if (activeJumping[a].name === rName && isPhoneMatch) {
            var inParts = activeJumping[a].inTime.split(":");
            var inTotalMin = (parseInt(inParts[0], 10) * 60) + parseInt(inParts[1], 10);
            var matchDiff = Math.abs(resTotalMin - inTotalMin);
            if (matchDiff <= 90) {
              activeJumping[a].therapyTime = rTimeShort;
              activeJumping[a].therapyRoom = rRoom;
            }
          }
        }
        for (var b=0; b<activeTherapy.length; b++) {
          var mPhoneT = formatPhoneSafely(activeTherapy[b].phone);
          var rPhoneT = formatPhoneSafely(resData[k][1]);
          var isPhoneMatchT = (rPhoneT === mPhoneT) || (rPhoneT.length >= 8 && mPhoneT.length >= 8 && rPhoneT.slice(-8) === mPhoneT.slice(-8));
 
          if (activeTherapy[b].name === rName && isPhoneMatchT) {
            var inPartsT = activeTherapy[b].inTime.split(":");
            var inTotalMinT = (parseInt(inPartsT[0], 10) * 60) + parseInt(inPartsT[1], 10);
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

/**
 * 🏰 [🏰 100% 안전 보관 레거시] 기존 getAdminDashboardData() 로직
 */
function getAdminDashboardData_Legacy() {
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
    
    // 1. 로그 업데이트 (단일 범위 setValues() 호출로 4배 이상 가속)
    logSheet.getRange(rowIdx, cols.classes + 1, 1, 4).setValues([[data.classes, data.timeLog, "귀가", timeStr]]);
    
    var phoneStr = String(logSheet.getRange(rowIdx, cols.phone + 1).getValue()).replace(/[^0-9]/g, "");
    var memberName = logSheet.getRange(rowIdx, cols.name + 1).getValue();
    
    // [v46.30] 웰니스 통합 점수제 연동: 퇴실 시 출석 기본점수 및 운동 강도점수 자동 하사
    var timeLogVal = parseFloat(data.timeLog) || 0;
    var reasonText = String(logSheet.getRange(rowIdx, cols.reason + 1).getValue() || "");
    var isCombo = (reasonText.indexOf("복합") !== -1);
    
    // (1) 출석 기본점수 (+20 EXP 상향) 기록 -> D열 (센터방문_수행)
    recordActivityLog({
      phone: phoneStr,
      name: memberName,
      type: "출석",
      item: "센터방문",
      action: "퇴실",
      score: 20,
      statType: "perf"
    });
    
    // (2) 운동 강도점수 (소수점 비례 반영: 운동시간 * 40 EXP) 기록
    if (timeLogVal > 0) {
      var extraWorkoutScore = Math.round(timeLogVal * 40);
      recordActivityLog({
        phone: phoneStr,
        name: memberName,
        type: "출석",
        item: "운동강도",
        action: "퇴실",
        score: extraWorkoutScore,
        statType: "perf"
      });
    }
    
    // (3) 테라피 회복력점수 (+30 EXP 상향) 기록 (0시간 테라피 전용이거나 복합출석인 경우 지급)
    if (timeLogVal === 0 || isCombo) {
      recordActivityLog({
        phone: phoneStr,
        name: memberName,
        type: "테라피",
        item: "스트레스지수감소",
        action: "완료",
        score: 30,
        statType: "def"
      });
    }

    // [v46.31] 주간 3회 출석 달성 시 체력 보너스 (+100 EXP) 자동 정산 연동
    checkAndAwardWeeklyAttendanceBonus(phoneStr, memberName);

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
    
    // 3. 테라피 예약 매칭 (완료 처리 - getValues() 및 초고속 안전 포맷터로 대폭 최적화)
    var resSheet = ss.getSheetByName("예약DB");
    var resData = resSheet.getDataRange().getValues();
    var memberName = logSheet.getRange(rowIdx, cols.name + 1).getValue();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    
    // 오늘 날짜 숫자로 단일화 (예: 20260521)
    var todayNum = todayStr.replace(/[^0-9]/g, "");
    
    for (var k = 1; k < resData.length; k++) {
      var rDateVal = resData[k][3]; // 예약DB 4번째 열 (Index 3: 날짜)
      if (!rDateVal) continue;
      
      var rDateNum = "";
      if (rDateVal instanceof Date) {
        rDateNum = Utilities.formatDate(rDateVal, "GMT+9", "yyyyMMdd");
      } else {
        var rDateRaw = String(rDateVal).trim();
        rDateNum = rDateRaw.replace(/[^0-9]/g, "");
        if (rDateNum.length > 8) rDateNum = rDateNum.substring(0, 8);
        if (rDateNum.length < 8) {
           var dParts = rDateRaw.split(/[^0-9]/).filter(Boolean);
           if (dParts.length >= 3) {
              rDateNum = dParts[0] + (dParts[1].length === 1 ? "0"+dParts[1] : dParts[1]) + (dParts[2].length === 1 ? "0"+dParts[2] : dParts[2]);
           }
        }
      }
      
      var rPhone = formatPhoneSafely(resData[k][1]); // Index 1: 전화번호
      var isPhoneMatch = (rPhone === phoneStr) || (rPhone.length >= 8 && phoneStr.length >= 8 && rPhone.slice(-8) === phoneStr.slice(-8));

      // 이름 매칭 생략: 날짜와 전화번호만으로 매칭
      if (rDateNum === todayNum && isPhoneMatch) {
        var status = String(resData[k][9] || ""); // Index 9: 상태
        if (status.indexOf("테라피") !== -1 || status.indexOf("예약") !== -1) {
          resSheet.getRange(k + 1, 10).setValue("귀가");
          resSheet.getRange(k + 1, 12).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm:ss"));
          break;
        }
      }
    }
    
    // 4. 퇴실 시 대기 문자 자동 취소 처리 (상태: 대기 -> 완료(출석))
    try {
      var smsSheet = ss.getSheetByName("문자발송");
      if (smsSheet) {
        var smsData = smsSheet.getDataRange().getDisplayValues();
        for (var j = 1; j < smsData.length; j++) {
          var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
          var sStatus = smsData[j][5];
          if (sPhone === phoneStr && sStatus === "대기") {
            smsSheet.getRange(j + 1, 6).setValue("완료(출석)");
          }
        }
      }
    } catch (smsErr) {
      Logger.log("체크아웃 대기 문자 취소 중 오류: " + smsErr.toString());
    }
    
    // 5. 글리코겐 방패 실시간 진도 계산 및 쪽지 발송 강제 유도
    try {
      getActiveQuestStatus(phoneStr, ss, null, memberName);
    } catch (qErr) {
      Logger.log("체크아웃 글리코겐 퀘스트 강제 업데이트 오류: " + qErr.toString());
    }
    
    return { success: true, message: "수정 및 퇴실 처리가 완료되었습니다." };
  } catch (e) {
    return { error: "수정 처리 오류: " + e.toString() };
  }
}

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

    // [캐시 갱신] 태블릿 PWA 출석기에서 실시간으로 연장된 정보가 노출되도록 구글 서버 캐시 강제 무효화
    try {
      var cache = CacheService.getScriptCache();
      cache.remove("v45_member_registry");
    } catch(cacheErr) {
      console.error("회원 레지스트리 캐시 삭제 실패: " + cacheErr.toString());
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
               "선택하신 [" + membership + "] 등록이 완료되었습니다.\n\n" +
               "📅 이용 기한: ~ " + expireDate + "까지\n" +
               "📊 이용 횟수: " + remainCount + "회\n\n" +
               "⚠️ 꼭 확인해 주세요!\n" +
               "소중한 회원권은 이용 기한인 **" + expireDate + "** 전까지 남은 " + remainCount + "회를 모두 소진하셔야 합니다. 늦지 않게 자주 방문하셔서 활기찬 하루를 충전해 보세요! 회원님의 건강과 아름다움을 위해 최선을 다해 돕겠습니다! 🥰";
  } else if (category === "연장결제") {
    content += "잊지 않고 연장 등록 해주셔서 감사합니다! ✨\n" +
               "기존 잔여분과 꼼꼼히 합산하여 [" + membership + "] 등록을 마쳤습니다.\n\n" +
               "📅 최종 이용 기한: ~ " + expireDate + "까지\n" +
               "📊 총 잔여 횟수: " + remainCount + "회\n\n" +
               "⚠️ 꼭 확인해 주세요!\n" +
               "이용 기한이 지나면 소중한 잔여 횟수가 소멸되오니, 최종 만료일인 **" + expireDate + "** 전까지 남은 " + remainCount + "회를 꼭 전부 사용해 주세요! 늦기 전에 서둘러 클럽에 나오셔서 신나게 점핑해 보아요! 꾸준한 관리가 최고의 결과를 만듭니다. 화이팅! 🔥";
  } else if (category === "재결제") {
    var greeting = (gapDays > 30) ? "다시 노형점핑을 찾아주셔서 정말 기뻐요! 🥰\n" : "노형점핑을 다시 믿고 선택해 주셔서 감사합니다! 😊\n";
    content += greeting +
               "결제하신 [" + membership + "] 등록이 완료되었습니다.\n\n" +
               "📅 이용 기한: ~ " + expireDate + "까지\n" +
               "📊 이용 횟수: " + remainCount + "회\n\n" +
               "⚠️ 꼭 확인해 주세요!\n" +
               "회원권 이용 기한인 **" + expireDate + "** 전까지 남은 " + remainCount + "회를 건강하게 소진해 주셔야 합니다. 소중한 운동 투자가 낭비되지 않도록 서둘러 내방해 주세요! 이번에도 회원님의 놀라운 변화를 위해 최선을 다하겠습니다! 💪";
  } else if (category === "추가결제") {
    content += "기존 프로그램과 더불어 [" + membership + "]을 추가해 주셔서 감사합니다! 🥰\n" +
               "회원님의 뜨거운 열정을 보며 저희도 큰 힘을 얻습니다.\n\n" +
               "📅 신규 이용 기한: ~ " + expireDate + "까지\n" +
               "📊 신규 이용 횟수: " + remainCount + "회\n\n" +
               "⚠️ 꼭 확인해 주세요!\n" +
               "새로 추가된 권종의 기한은 **" + expireDate + "**까지입니다! 꼭 기한 내에 남은 " + remainCount + "회를 시너지 나게 모두 즐기실 수 있도록 자주 방문해 주세요! 두 가지 모두 빛을 발하게 세심하게 케어해 드릴게요! ✨";
  } else {
    content += "결제하신 [" + membership + "] 등록이 완료되었습니다.\n\n" +
               "📅 이용 기한: ~ " + expireDate + "까지\n" +
               "📊 이용 횟수: " + remainCount + "회\n\n" +
               "⚠️ 꼭 확인해 주세요!\n" +
               "이용 기한인 **" + expireDate + "** 전까지 남은 " + remainCount + "회를 모두 소진해 주셔야 하니 늦지 않게 자주 방문해 주세요! 언제나 정성을 다해 관리해 드리겠습니다! 💪";
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

function updateWorkLog(data) {
  return submitWorkLog(data);
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

function updateSalesRecord(data) {
  try {
    if (!data.id) return { error: "매출 고유 ID가 누락되었습니다." };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName("판매내역");
    if (!salesSheet) return { error: "판매내역 시트가 없습니다." };
    
    var values = salesSheet.getDataRange().getValues();
    var foundRow = -1;
    
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(data.id).trim()) {
        foundRow = i + 1; // row index (1-based)
        break;
      }
    }
    
    if (foundRow === -1) {
      return { error: "수정하려는 매출 내역을 찾을 수 없습니다." };
    }
    
    // B열 날짜 셀 저장을 위해 시간 추출 및 보존
    var oldDateTime = values[foundRow - 1][1]; 
    var timeStr = "00:00:00";
    if (oldDateTime instanceof Date) {
      timeStr = Utilities.formatDate(oldDateTime, "GMT+9", "HH:mm:ss");
    } else {
      var parts = String(oldDateTime).split(" ");
      if (parts.length > 1) {
        timeStr = parts[1];
      }
    }
    
    var finalDateStr = data.date + " " + timeStr;
    
    // B열(날짜), C열(구분), D열(지급처/구입자), E열(세부항목명), F열(금액), G열(결제수단), H열(비고) 순차 업데이트
    salesSheet.getRange(foundRow, 2).setValue(finalDateStr);
    salesSheet.getRange(foundRow, 3).setValue(data.category);
    salesSheet.getRange(foundRow, 4).setValue(data.buyer);
    salesSheet.getRange(foundRow, 5).setValue(data.itemName || data.category);
    salesSheet.getRange(foundRow, 6).setValue(Number(data.amount) || 0);
    salesSheet.getRange(foundRow, 7).setValue(data.payMethod);
    salesSheet.getRange(foundRow, 8).setValue(data.memo || "");
    
    return { success: true, message: "매출 기록이 성공적으로 수정되었습니다." };
  } catch (e) {
    return { error: "매출 수정 오류: " + e.toString() };
  }
}

function deleteSalesRecord(id) {
  try {
    if (!id) return { error: "삭제하려는 고유 ID가 누락되었습니다." };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName("판매내역");
    if (!salesSheet) return { error: "판매내역 시트가 없습니다." };
    
    var values = salesSheet.getDataRange().getValues();
    var foundRow = -1;
    
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === String(id).trim()) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow === -1) {
      return { error: "삭제하려는 매출 내역을 찾을 수 없습니다." };
    }
    
    salesSheet.deleteRow(foundRow);
    return { success: true, message: "매출 기록이 성공적으로 삭제되었습니다." };
  } catch (e) {
    return { error: "매출 삭제 오류: " + e.toString() };
  }
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
 * 33 챌린지 전체 기록 조회 ('33챌린지_인바디' 시트 기준 - 관리자 고도화)
 */
function getInBodyHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지_인바디");
    if (!sheet) return { success: true, records: [] };
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, records: [] };
    
    var records = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dateVal = row[0]; // A열: 측정 날짜
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, "GMT+9", "yyyy-MM-dd");
      }
      
      var wScore = row[6] !== "" ? Number(row[6]) : 0; // G열: 주간점수
      var mScore = row[7] !== "" ? Number(row[7]) : 0; // H열: 월간점수
      var tScore = row[8] !== "" ? Number(row[8]) : 0; // I열: 토탈점수
      
      records.push({
        rowIdx: i + 1,
        date: String(dateVal || ""),
        name: String(row[1] || ""), // B열: 이름
        phone: String(row[2] || ""), // C열: 연락처
        weight: row[3] !== "" ? Number(row[3]) : 0, // D열: 체중
        muscle: row[4] !== "" ? Number(row[4]) : 0, // E열: 골격근량
        fat: row[5] !== "" ? Number(row[5]) : 0, // F열: 체지방률
        weeklyScore: wScore,
        monthlyScore: mScore,
        totalScore: tScore,
        score: wScore + mScore + tScore, // 합산점수
        memo: String(row[9] || "") // J열: 비고
      });
    }
    // 최신순 정렬
    records.reverse();
    return { success: true, records: records };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * 33 챌린지 인바디 기록 개별 수정 API
 */
function updateInBodyRecord(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("33챌린지_인바디");
    if (!sheet) return { success: false, error: "시트를 찾을 수 없습니다." };
    
    var rowIdx = Number(payload.rowIdx);
    if (!rowIdx || rowIdx < 2) return { success: false, error: "유효하지 않은 행 인덱스입니다." };
    
    var name = String(payload.name).trim();
    var phone = formatPhoneNumber(payload.phone).replace(/[^0-9]/g, "");
    var weight = Number(payload.weight || 0);
    var muscle = Number(payload.muscle || 0);
    var fat = Number(payload.fat || 0);
    var memo = String(payload.memo || "").trim();

    // 회원명단 시트에서 목표체중 로드
    var targetWeight = 0;
    var memberSheet = ss.getSheetByName("회원명단");
    if (memberSheet) {
      var mData = memberSheet.getDataRange().getDisplayValues();
      var mCols = getMemberSheetColumnIndices(memberSheet);
      for (var i = 1; i < mData.length; i++) {
        var mPhone = normalizePhoneDigits(mData[i][mCols.phone]);
        if (mPhone === phone || (phone.length >= 8 && mPhone.endsWith(phone.substring(phone.length - 8)))) {
          targetWeight = parseWeightSafely(mData[i][mCols.targetWeight]);
          break;
        }
      }
    }
    
    var targetDate = new Date(payload.date);
    if (isNaN(targetDate.getTime())) {
      return { success: false, error: "유효하지 않은 날짜 포맷입니다." };
    }
    
    // 1. 전체 이력 조회 및 이번에 수정할 행을 반영한 가상 목록 구성
    var data = sheet.getDataRange().getValues();
    var userRecords = [];
    
    for (var i = 1; i < data.length; i++) {
      var rowPhone = formatPhoneNumber(data[i][2]).replace(/[^0-9]/g, "");
      if (rowPhone === phone) {
        var rDate = new Date(data[i][0]);
        if (isNaN(rDate.getTime())) continue;
        
        // 현재 수정 대상인 행인 경우 새 수치로 대체, 다른 행은 시트의 수치 사용
        if (i + 1 === rowIdx) {
          userRecords.push({
            date: targetDate,
            weight: weight,
            muscle: muscle,
            fat: fat,
            rowIdx: i + 1
          });
        } else {
          userRecords.push({
            date: rDate,
            weight: Number(data[i][3] || 0),
            muscle: Number(data[i][4] || 0),
            fat: Number(data[i][5] || 0),
            rowIdx: i + 1
          });
        }
      }
    }
    
    // 만약 현재 수정 중인 행이 userRecords에 없다면 추가
    var foundCurrent = false;
    for (var rIdx = 0; rIdx < userRecords.length; rIdx++) {
      if (userRecords[rIdx].rowIdx === rowIdx) {
        foundCurrent = true;
        break;
      }
    }
    if (!foundCurrent) {
      userRecords.push({
        date: targetDate,
        weight: weight,
        muscle: muscle,
        fat: fat,
        rowIdx: rowIdx
      });
    }

    // 날짜 순 정렬 (오름차순)하여 3단 스탯 변화를 타임라인으로 정밀 연산
    userRecords.sort(function(a, b) { return a.date - b.date; });

    // 인바디 계산 공통 헬퍼
    function calculateInbodyScoreHelper(first, current, targetWeight, scoreType) {
      if (!first || !current) return 0;
      var score = 0;
      var fW = Number(first.weight) || 0;
      var cW = Number(current.weight) || 0;
      var fM = Number(first.muscle) || 0;
      var cM = Number(current.muscle) || 0;
      var fF = Number(first.fat) || 0;
      var cF = Number(current.fat) || 0;
      
      var diffW = Number((fW - cW).toFixed(2)) || 0;
      var diffM = Number((cM - fM).toFixed(2)) || 0;
      var diffFat = Number((fF - cF).toFixed(1)) || 0;
      
      if (diffW > 0) score += (diffW * 10) * 50;
      if (diffM > 0) score += (diffM * 10) * 200;
      if (diffFat > 0) score += (diffFat * 10) * 100;
      
      var firstFatMass = fW * (fF / 100);
      var currentFatMass = cW * (cF / 100);
      var fatLossRate = 0;
      if (firstFatMass > 0 && !isNaN(firstFatMass)) {
        fatLossRate = ((firstFatMass - currentFatMass) / firstFatMass) * 100;
      }
      if (diffW >= 8.0 || (fatLossRate >= 20.0 && !isNaN(fatLossRate))) {
        score += 1500;
      }
      
      // 🏆 체성분 명품 유지 보너스 판정 엔진 (±0.5kg 정교화)
      if (targetWeight && targetWeight > 0) {
        if (scoreType === "monthly") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        } else if (scoreType === "lifetime") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        }
      }
      
      return isNaN(score) ? 0 : score;
    }

    var targetRecord = null;
    var targetIdxInList = -1;
    for (var rIdx = 0; rIdx < userRecords.length; rIdx++) {
      if (userRecords[rIdx].rowIdx === rowIdx) {
        targetRecord = userRecords[rIdx];
        targetIdxInList = rIdx;
        break;
      }
    }

    // 날짜 기준점 계산
    var day = targetRecord.date.getDay();
    var diffToThu = (day + 3) % 7; 
    var startOfWeek = new Date(targetRecord.date.getFullYear(), targetRecord.date.getMonth(), targetRecord.date.getDate() - diffToThu);
    startOfWeek.setHours(0, 0, 0, 0);
    
    var year = targetRecord.date.getFullYear();
    var month = targetRecord.date.getMonth();
    var startOfMonth = new Date(year, month, 1, 0, 0, 0, 0);
    
    var firstDayOfWeek = startOfMonth.getDay();
    var diffToWed = (3 - firstDayOfWeek + 7) % 7;
    var baselineDeadline = new Date(year, month, 1 + diffToWed);
    if (typeof isCenterHoliday === "function" && isCenterHoliday(baselineDeadline)) {
      baselineDeadline.setDate(baselineDeadline.getDate() + 1);
    }
    baselineDeadline.setHours(23, 59, 59, 999);

    var firstEver = null;
    var prevBeforeThisWeek = null;
    var baselineRecord = null;

    for (var j = 0; j <= targetIdxInList; j++) {
      var r = userRecords[j];
      
      if (!firstEver || r.date < firstEver.date) {
        firstEver = r;
      }
      
      if (r.date < startOfWeek) {
        var timeDiff = startOfWeek.getTime() - r.date.getTime();
        var diffDays = timeDiff / (1000 * 60 * 60 * 24);
        if (diffDays <= 7.5) {
          if (!prevBeforeThisWeek || r.date > prevBeforeThisWeek.date) {
            prevBeforeThisWeek = r;
          }
        }
      }
      
      if (r.date >= startOfMonth && r.date <= baselineDeadline) {
        if (!baselineRecord || r.date < baselineRecord.date) {
          baselineRecord = r;
        }
      }
    }

    var weeklyScore = 0;
    var monthlyScore = 0;
    var totalScore = 0;

    if (firstEver && firstEver !== targetRecord) {
      totalScore = calculateInbodyScoreHelper(firstEver, targetRecord, targetWeight, 'lifetime');
    }
    if (prevBeforeThisWeek) {
      weeklyScore = calculateInbodyScoreHelper(prevBeforeThisWeek, targetRecord, targetWeight, 'weekly');
    }
    if (baselineRecord && baselineRecord !== targetRecord) {
      if (isDateInLastWeekMonToWed(targetRecord.date)) {
        monthlyScore = calculateInbodyScoreHelper(baselineRecord, targetRecord, targetWeight, 'monthly');
      } else {
        monthlyScore = 0;
      }
    }
    
    // 시트 셀에 하나씩 매칭하여 저장 (11열 포맷)
    sheet.getRange(rowIdx, 1).setValue(targetDate);
    sheet.getRange(rowIdx, 2).setValue(name);
    sheet.getRange(rowIdx, 3).setValue("'" + phone);
    sheet.getRange(rowIdx, 4).setValue(weight);
    sheet.getRange(rowIdx, 5).setValue(muscle);
    sheet.getRange(rowIdx, 6).setValue(fat);
    sheet.getRange(rowIdx, 7).setValue(weeklyScore);
    sheet.getRange(rowIdx, 8).setValue(monthlyScore);
    sheet.getRange(rowIdx, 9).setValue(totalScore);
    sheet.getRange(rowIdx, 10).setValue(memo);
    
    // [perf] 인바디 업로드 완료 시 대시보드 캐시 즉시 파괴
    clearUserDashboardCache(phone);
    
    return { success: true, message: "기록이 수정되었으며 점수가 자동 재계산되었습니다." };
  } catch (e) {
    return { success: false, error: e.toString() };
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
 */
// [v63.0] 레거시 함수 제거 완료

/**
 * 노형점핑클럽 - 전체 안내 및 잔여 현황 발송 팝업
 */
// [v63.0] 레거시 함수 제거 완료

function setCheck(r) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("등록 현황");
  sheet.getRange(r, 14).setValue(true); 
}

function setAllChecks(rs) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("등록 현황");
  rs.forEach(function(r) { sheet.getRange(r, 14).setValue(true); });
}

// [v63.0] 레거시 함수 제거 완료

/**
 * 밤 11시 30분경 자동으로 실행되어 그날의 최종 통계를 업무일지에 업데이트합니다.
 * (트리거 설정 필요: autoCloseDailyLog / 시간 기반 / 일일 타이머 / 오후 11시~자정)
 */
function autoCloseDailyLog() {
  try {
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 0. [v60.0] 휴무일 및 공휴일 감지 (자동 마감 스킵 정책)
    var dayOfWeek = now.getDay(); // 0: 일요일, 6: 토요일
    var isOff = (dayOfWeek === 0); // 일요일 기본 휴무
    
    // 법정 공휴일 체크
    if (!isOff && isKoreanPublicHoliday(now)) {
      isOff = true;
    }
    
    // '벙개테라피 및 휴일 설정' 시트 체크
    var isFlash = false;
    try {
      var flashSheet = ss.getSheetByName('벙개테라피 및 휴일 설정');
      if (flashSheet) {
        var flashData = flashSheet.getDataRange().getDisplayValues().slice(1);
        for (var f = 0; f < flashData.length; f++) {
          var fDate = String(flashData[f][0]);
          if (fDate === dateStr) {
            if (String(flashData[f][1]) === "휴무") {
              isOff = true;
            }
            if (String(flashData[f][2]) === "벙개") {
              isFlash = true;
              isOff = false; // 벙개면 휴무라도 영업함
            }
            break;
          }
        }
      }
    } catch (e) {
      Logger.log("휴일 설정 시트 조회 오류: " + e.toString());
    }
    
    if (isOff && !isFlash) {
      Logger.log(dateStr + "은 휴무일이므로 자동 마감 일지 작성을 스킵합니다.");
      return { success: true, message: "휴무일 마감 스킵 완료" };
    }
    
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
    
    // 2. 최종 통계 가져오기
    var statsRes = getTodayWorkLogStats(dateStr);
    if (statsRes.error) {
      Logger.log("자동 마감 통계 조회 실패: " + statsRes.error);
      return;
    }
    var stats = statsRes.stats;
    
    // 3. 기존 일지가 있는지 확인
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
    
    // 4. 전송용 데이터 조립 (기존 내용 보존하면서 실시간 강제 통계 업데이트 구현)
    var data = {
      date: dateStr,
      author: existingLog ? String(existingLog[1] || "") : "시스템 자동마감",
      jumpingList: existingLog ? String(existingLog[2] || "") : "(미기입)",
      muscleList: existingLog ? String(existingLog[3] || "") : "(미기입)",
      stats: stats,
      remarks: existingLog ? String(existingLog[13] || "") : "퇴근 전 미작성되어 시스템에 의해 자동 마감되었습니다.",
      issues: existingLog ? String(existingLog[14] || "") : ""
    };
    
    // 5. 저장 함수 호출 (언제나 최신 통계로 덮어쓰기 완료)
    var res = submitWorkLog(data);
    Logger.log(dateStr + " 자동 마감 결과: " + res.message);
    return res;
    
  } catch (e) {
    Logger.log("자동 마감 치명적 오류: " + e.toString());
    return { error: e.toString() };
  }
}

// [v63.0] 레거시 함수 제거 완료

// [v63.0] 레거시 함수 제거 완료

// [v63.0] backupRemainingSessions 레거시 함수 제거 완료

// [v63.0] 레거시 함수 제거 완료
// [v63.0] 레거시 함수 제거 완료

// [v63.0] 레거시 함수 제거 완료

// [v63.0] 레거시 함수 제거 완료

// [v63.0] 레거시 함수 제거 완료
// [v63.0] 레거시 함수 제거 완료

// [v63.0] 레거시 함수 제거 완료
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
 * 🌿 [v64.80] 관리자용 특정 기간 범위 (startDateStr로부터 durationDays 동안) 테라피 예약 조회 API
 */
function getAdminReservationDataRange(startDateStr, durationDays) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("예약DB");
    if (!sheet) return [];
    
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

    var limitDays = Number(durationDays) || 7;
    var targetDates = [];
    var startD = new Date(startDateStr);
    for (var d = 0; d < limitDays; d++) {
      var temp = new Date(startD);
      temp.setDate(startD.getDate() + d);
      var y = temp.getFullYear();
      var m = String(temp.getMonth() + 1).padStart(2, '0');
      var dateDay = String(temp.getDate()).padStart(2, '0');
      targetDates.push(y + "-" + m + "-" + dateDay);
    }
    
    var results = [];
    for (var i = 1; i < data.length; i++) {
      var rDateStr = String(data[i][idx.date] || "").split(" ")[0]; // YYYY-MM-DD
      if (!rDateStr) continue;
      
      if (targetDates.indexOf(rDateStr) !== -1) {
        var rawTime = data[i][idx.time];
        var timeStr = "시간미정";
        
        if (rawTime && rawTime.includes(':')) {
          var parts = rawTime.split(':');
          if (parts.length >= 2) {
            var hh = parts[0].trim().padStart(2, '0');
            var mm = parts[1].trim().padStart(2, '0');
            timeStr = hh + ":" + mm;
          }
        }
        timeStr = timeStr.replace(/[^0-9:]+$/, "");

        results.push({
          rowIdx: i + 1,
          date: rDateStr,
          time: timeStr, 
          name: data[i][idx.name],
          room: data[i][idx.room],
          people: data[i][idx.people],
          status: data[i][idx.status]
        });
      }
    }
    
    // 날짜순 -> 시간순 -> 방이름 순 정렬
    results.sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      if (a.time !== b.time) return a.time.localeCompare(b.time);
      return a.room.localeCompare(b.room);
    });
    
    return results;
  } catch (e) {
    return [];
  }
}

/**
 * 🌿 [v64.80] 관리자용 기존 테라피 예약 정보 정밀 수정 API
 */
function updateReservation(rowIdx, payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('예약DB');
    if (!sheet) return { error: "예약DB 시트를 찾을 수 없습니다." };
    
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    
    var idx = { date: 3, time: 4, start: 5, end: 6, room: 7 }; 
    for (var h = 0; h < headers.length; h++) {
      var head = String(headers[h]).trim();
      if (head.indexOf("예약날짜") !== -1 || head.indexOf("날짜") !== -1) idx.date = h;
      if (head.indexOf("입실시간") !== -1 || head.indexOf("예약시간") !== -1) idx.time = h;
      if (head.indexOf("사우나시작") !== -1) idx.start = h;
      if (head.indexOf("사우나종료") !== -1) idx.end = h;
      if (head.indexOf("배정방") !== -1 || head.indexOf("방") !== -1) idx.room = h;
    }
    
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

    var startTime = formatTime(payload.time); 
    var saunaStart = addMinutes(payload.time, 30);
    var saunaEnd = addMinutes(payload.time, 80);
    
    // 예약DB의 rowIdx 행에 있는 개별 컬럼 셀들 수정
    sheet.getRange(rowIdx, idx.date + 1).setValue(payload.date);
    sheet.getRange(rowIdx, idx.time + 1).setValue(startTime);
    sheet.getRange(rowIdx, idx.start + 1).setValue(saunaStart);
    sheet.getRange(rowIdx, idx.end + 1).setValue(saunaEnd);
    sheet.getRange(rowIdx, idx.room + 1).setValue(payload.room);
    
    SpreadsheetApp.flush();
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
 * 중복 방지를 위해 밤 11시 30분 자동 마감 함수(autoCloseDailyLog) 구버전 블록을 완전 삭제 및 병합 조치했습니다.
 * 실제 자동 퇴실 처리가 포함된 최신버전의 autoCloseDailyLog는 4274번 줄에 완벽하게 살아있습니다!
 */
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
    // 자동 생성 대기열은 각각의 안부/결석/디버프 스캔 함수 진입 시점에 안전하게 초기 청소 후 최신 정보로 재생성됩니다.
    
    // 다시 추출


    checkLongTermAbsentees();         // 💌 [자동화] 7일 이상 장기 미방문자 안부 문자 자동 스캔 적재!
    checkInactiveMembers();          // 🔄 [자동화] 14일 이상 장기 미등록자 복귀 권유 문자 자동 스캔 적재!
    checkInactivityDebuffAbsentees(); // [v46.35] 연속 미출석 결석 디버프 대기열 추가
    
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
 * [관리자 전용] 문자발송 내용 및 상태 동시 업데이트 (완료 등)
 */
function updateSmsContentAndStatus(rowIdx, content, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("문자발송");
    if (!sheet) return { error: "시트를 찾을 수 없습니다." };
    
    sheet.getRange(rowIdx, 5).setValue(content); // 5열: 생성된문자내용
    sheet.getRange(rowIdx, 6).setValue(status);  // 6열: 상태
    return { success: true, message: "내용 수정 및 상태가 " + status + "(으)로 안전하게 업데이트되었습니다." };
  } catch (e) {
    return { error: "수정 및 업데이트 실패: " + e.toString() };
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
    
    // [중복 업데이트 및 실시간 정밀 타임라인 덮어쓰기 알고리즘]
    // 1. 기존에 시트에 존재하는 모든 '대기' 상태의 '장기미방문' 문자열 행들을 스캔하여 완벽히 삭제(Clean-up)합니다!
    var smsDataForCleanup = smsSheet.getDataRange().getValues();
    for (var i = smsDataForCleanup.length - 1; i >= 1; i--) {
      var category = String(smsDataForCleanup[i][3]);
      var status = String(smsDataForCleanup[i][5]).trim();
      if (category === "장기미방문" && status === "대기") {
        smsSheet.deleteRow(i + 1);
      }
    }
    
    // 삭제 처리 완료 후 최신 데이터 로드!
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
    
    // 2. 이미 발송 대기 중이거나 완료된 목록 확인 (중복 방지 및 7일 주기 재발송용)
    var hasPendingSmsMap = {};
    var lastSentSmsDateMap = {};
    for (var j = 1; j < smsData.length; j++) {
      var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
      var sCategory = smsData[j][3];
      var sStatus = smsData[j][5];
      var sDateStr = smsData[j][0]; // "기록시간"
      
      if (sCategory === "장기미방문") {
        if (sStatus === "대기") {
          hasPendingSmsMap[sPhone] = true;
        } else if (sStatus === "완료") {
          var sDate = new Date(sDateStr);
          if (!lastSentSmsDateMap[sPhone] || sDate > lastSentSmsDateMap[sPhone]) {
            lastSentSmsDateMap[sPhone] = sDate;
          }
        }
      }
    }

    // 2-2. [v64.60] 회원별 보유 회원권 목록 및 출석 횟수 사전 집계 (정밀 성향 분석용)
    var memberMembershipNamesMap = {}; // { phone: [membership1, membership2, ...] }
    for (var k = 1; k < regData.length; k++) {
      var status = String(regData[k][regCols.status]).trim();
      if (status !== "진행중" && status !== "진행 중") continue;
      var phone = String(regData[k][regCols.phone] || "").replace(/[^0-9]/g, "");
      var membership = String(regData[k][regCols.membership] || "").trim();
      if (!phone || !membership) continue;
      
      if (!memberMembershipNamesMap[phone]) {
        memberMembershipNamesMap[phone] = [];
      }
      if (memberMembershipNamesMap[phone].indexOf(membership) === -1) {
        memberMembershipNamesMap[phone].push(membership);
      }
    }

    var memberAttendanceTypeMap = {}; // { phone: { therapy: 0, jumping: 0 } }
    for (var l = 1; l < logData.length; l++) {
      var lPhone = String(logData[l][logCols.phone] || "").replace(/[^0-9]/g, "");
      var lType = String(logData[l][logCols.type] || ""); // 출석유형
      if (!lPhone) continue;
      
      if (!memberAttendanceTypeMap[lPhone]) {
        memberAttendanceTypeMap[lPhone] = { therapy: 0, jumping: 0 };
      }
      if (lType.indexOf("테라피") !== -1 || lType.indexOf("원적외선") !== -1 || lType.indexOf("반신욕") !== -1) {
        memberAttendanceTypeMap[lPhone].therapy++;
      } else {
        memberAttendanceTypeMap[lPhone].jumping++;
      }
    }
    
    // 3. 미방문자 추출 및 문자/쪽지 생성 (양동작전)
    var count = 0;
    var addedNames = [];
    
    for (var k = 1; k < regData.length; k++) {
      var status = String(regData[k][regCols.status]).trim();
      if (status !== "진행중" && status !== "진행 중") continue;
      
      var phone = String(regData[k][regCols.phone] || "").replace(/[^0-9]/g, "");
      var name = regData[k][regCols.name];
      if (!phone) continue;
      
      var lastDate = lastAttendanceMap[phone];
      var cleanName = name.replace(/\d{4}$/, ""); // 이름 뒤 번호 제거
      
      // 마지막 출석일이 없거나(한번도 안옴), 7일 이상 지난 경우
      if (!lastDate || lastDate < sevenDaysAgo) {
        // 이미 대기 중인 문자가 있으면 패스
        if (hasPendingSmsMap[phone]) continue;
        
        // 마지막 발송 정보 검사
        var lastSentSmsDate = lastSentSmsDateMap[phone];
        if (lastSentSmsDate) {
          // 출석 기록이 마지막 발송일보다 최신인 경우 -> 정상 재생성 대상
          // 출석 기록이 마지막 발송일보다 이전인 경우 -> 7일간 미방문이 지속되었는지 검사
          if (lastDate && lastDate <= lastSentSmsDate) {
            var diffMs = now.getTime() - lastSentSmsDate.getTime();
            var daysSinceLastSms = diffMs / (1000 * 60 * 60 * 24);
            if (daysSinceLastSms < 7) {
              continue; // 아직 7일이 지나지 않았으므로 건너뜀 (스팸 방지)
            }
          }
        }
        
        // 결석 일수 계산
        var absentDays = 999; // 기록 없음
        if (lastDate) {
          absentDays = Math.floor((now.getTime() - lastDate.getTime()) / (1000 * 60 * 60 * 24));
        }
        
        // 회원권 정보 추출
        var membership = regData[k][regCols.membership] || "회원권";
        var expire = regData[k][regCols.expire] || "미정";
        var remain = regData[k][regCols.remain] || "0";
        var expWarn = "\n🎫 회원님의 [" + membership + "] 만료일은 " + expire + "까지이며, 현재 잔여 횟수가 " + remain + "회 남아있습니다. 소중한 이용권이 마감되어 소멸되기 전에 꼭 오셔서 알차게 사용하시길 바랄게요! 😊";
        
        var formattedPhone = formatPhoneForSms(regData[k][regCols.phone]);
        
        // 🚨 [v64.60] 장기미방문 결석 회원도 복합 회원권 및 다중 보유 성향 분석 반영 (공용 헬퍼 적용)
        var mTypes = memberMembershipNamesMap[phone] || [membership];
        var attStats = memberAttendanceTypeMap[phone] || { therapy: 0, jumping: 0 };
        var regClass = determineMemberClassInfo(mTypes, attStats);
        
        // 🤖 [v64.50] 제미나이 AI 기반 히스토리-성향-맥락 융합 본문 2채널 동적 생성 (만료일 인자 수혈하여 버그 100% 박멸!)
        var msg = generateWellnessAiSms(cleanName, remain, absentDays, "장기미방문", formattedPhone, "sms", regClass, false, expire);
        
        var notiTitle = "회원님, 에너지를 다시 채워드릴게요! ❤️";
        if (absentDays >= 14 && absentDays <= 29) {
          notiTitle = "회원님의 활기찬 에너지가 무척 그립습니다! 😊";
        } else if (absentDays >= 30) {
          notiTitle = "회원님, 건강한 습관을 항상 응원합니다! 🌟";
        }
        
        var notiContent = generateWellnessAiSms(cleanName, remain, absentDays, "장기미방문", formattedPhone, "noti", regClass, false, expire) + expWarn;
        
        // A. 오프라인 문자 적재
        smsSheet.appendRow([
          Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm"),
          cleanName,
          formattedPhone,
          "장기미방문",
          msg,
          "대기"
        ]);
        
        // B. 인앱 쪽지 실시간 동시 전송 (양동작전)
        try {
          if (!hasSentNotificationToday(phone, "안부", notiTitle)) {
            sendPersonalNotification(phone, "안부", notiTitle, notiContent);
          }
        } catch (notiErr) {
          Logger.log("장기미방문 인앱 쪽지 생성 실패 (" + cleanName + "): " + notiErr.toString());
        }
        
        count++;
        addedNames.push(cleanName);
      }
    }
    
    // 4. 진행 중인 글리코겐 퀘스트 미출석 복구 독려 검사 엔진
    try {
      var questSheet = ss.getSheetByName("돌발퀘스트_목록");
      if (questSheet) {
        var questData = questSheet.getDataRange().getDisplayValues();
        var todayMidnight = new Date();
        todayMidnight.setHours(0,0,0,0);
        
        var localLogData = logData; // logData 공유
        
        for (var q = 1; q < questData.length; q++) {
          var qType = questData[q][2];
          var qStatus = questData[q][7];
          
          if (qType === "글리코겐" && qStatus === "진행중") {
            var qPhone = String(questData[q][6]).replace(/[^0-9]/g, "");
            var qStartStr = questData[q][1];
            
            // 경과 일수 계산 (시행일 기준)
            var qStartDate = new Date(qStartStr);
            qStartDate.setHours(0,0,0,0);
            var elapsedDays = Math.floor((todayMidnight.getTime() - qStartDate.getTime()) / (1000 * 60 * 60 * 24));
            
            // 현재 출석 진척도 획득
            var qStatusObj = getActiveQuestStatus(qPhone, ss, localLogData);
            if (qStatusObj && qStatusObj.glycogenQuest) {
              var qProgress = qStatusObj.glycogenQuest.progress;
              
              if (elapsedDays === 1 && qProgress === 0) {
                // 1일차 미출석 경고 쪽지 발송
                if (!hasSentNotificationToday(qPhone, "방어", "1일차 경고")) {
                  sendPersonalNotification(
                    qPhone,
                    "방어",
                    "🚨 글리코겐 방패 1일차 경고! 내일부터 2배 빡세게! 🔥",
                    "회원님, 어제 출석을 못 하셨네요! 😢 글리코겐이 몸의 지방으로 굳어버리기 전에 남은 이틀 동안 2배 더 빡세게 운동하셔야 합니다! 내일은 꼭 클럽에 나오셔서 활기차게 뛰고 글리코겐 방패를 채워보세요. 웰니스 코치가 기다리고 있습니다! 화이팅! 🏃‍♀️"
                  );
                }
              } else if (elapsedDays === 2 && qProgress <= 1) {
                // 2일차 미출석 비상 상황 쪽지 발송
                if (!hasSentNotificationToday(qPhone, "방어", "비상 상황")) {
                  sendPersonalNotification(
                    qPhone,
                    "방어",
                    "🚨 글리코겐 방패 비상 상황! 마지막 조각 충전 기회! 🛡️",
                    "비상 상황입니다, 회원님! 🚨 벌써 글리코겐 방패 제한 시간의 마지막 날입니다. 아직 방패가 완성되지 않았어요! 오늘이 지방 축적을 막을 수 있는 마지막 골든타임입니다. 오늘 꼭 클럽에 오셔서 운동을 완수하고 요요 방어 방패를 꼭 완성하세요! 🔥"
                  );
                }
              }
            }
          }
        }
      }
    } catch (qErr) {
      Logger.log("글리코겐 미출석 복구 독려 스캔 중 에러: " + qErr.toString());
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
 * [자동화/관리자 공용] 등록이 끊긴 지 1일 이상 된 미등록 회원 스캔 및 복귀 권유 문자 자동 적재 (v64.20)
 * 🚨 [v66.1] 자동 낭비 방지를 위한 수동화 패치 완료
 */
function checkInactiveMembers(baseDateStr, isManual) {
  try {
    if (!isManual) {
      Logger.log("⚠️ 장기 미등록(복귀권유) 자동 생성 차단 (어드민 수동 생성만 허용)");
      return { success: false, error: "장기 미등록 문자 자동 생성은 멈춤 상태입니다. 어드민에서 수동으로 생성해 주세요." };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황");
    var smsSheet = ss.getSheetByName("문자발송");
    var logSheet = ss.getSheetByName("출석기록");
    
    if (!regSheet || !smsSheet) return { error: "필요한 시트(등록현황/문자발송)가 없습니다." };
    
    // [중복 업데이트 및 실시간 정밀 타임라인 덮어쓰기 알고리즘]
    // 1. 기존에 시트에 존재하는 모든 '대기' 상태의 '복귀권유' 문자열 행들을 스캔하여 완벽히 삭제(Clean-up)합니다!
    var smsDataForCleanup = smsSheet.getDataRange().getValues();
    for (var i = smsDataForCleanup.length - 1; i >= 1; i--) {
      var category = String(smsDataForCleanup[i][3]);
      var status = String(smsDataForCleanup[i][5]).trim();
      if (category === "복귀권유" && status === "대기") {
        smsSheet.deleteRow(i + 1);
      }
    }
    
    // 삭제 처리 완료 후 최신 데이터 로드!
    var regData = regSheet.getDataRange().getDisplayValues();
    var smsData = smsSheet.getDataRange().getDisplayValues();
    var logData = logSheet ? logSheet.getDataRange().getDisplayValues() : [];
    
    var regCols = getRegColumnIndices(regSheet);
    var logCols = logSheet ? getAttendanceColumnIndices(logSheet) : null;
    
    var now = baseDateStr ? new Date(baseDateStr) : new Date();
    now.setHours(23, 59, 59, 999); // 일수 계산 정밀성 극대화 및 기준일 연동
    
    // 1. 회원별 출석기록에서 테라피/점핑 횟수 집계 (출석 성향 분석용)
    var memberAttendanceTypeMap = {}; // { phone: { therapy: 0, jumping: 0 } }
    if (logSheet && logData.length >= 2 && logCols) {
      for (var l = 1; l < logData.length; l++) {
        var lPhone = String(logData[l][logCols.phone] || "").replace(/[^0-9]/g, "");
        var lType = String(logData[l][logCols.type] || ""); // 출석유형
        if (!lPhone) continue;
        
        if (!memberAttendanceTypeMap[lPhone]) {
          memberAttendanceTypeMap[lPhone] = { therapy: 0, jumping: 0 };
        }
        
        if (lType.indexOf("테라피") !== -1 || lType.indexOf("보너스") !== -1) {
          memberAttendanceTypeMap[lPhone].therapy++;
        } else if (lType.indexOf("점핑") !== -1 || lType.indexOf("운동") !== -1 || lType.indexOf("기본") !== -1) {
          memberAttendanceTypeMap[lPhone].jumping++;
        }
      }
    }
    
    // 2. 회원별 상태 체크 (가장 최근 만료일, 활성 여부 및 회원권 종류 취합)
    var memberInfoMap = {};
    for (var i = 1; i < regData.length; i++) {
      var phone = String(regData[i][regCols.phone] || "").replace(/[^0-9]/g, "");
      var status = String(regData[i][regCols.status] || "").trim(); 
      var expireDateStr = regData[i][regCols.expire];
      var name = regData[i][regCols.name];
      var remain = parseInt(regData[i][regCols.remain]) || 0;
      var membership = String(regData[i][regCols.membership] || "");
      
      if (!phone) continue;
      
      if (!memberInfoMap[phone]) {
        memberInfoMap[phone] = { 
          isActive: false, 
          lastExpire: new Date(0), 
          name: name, 
          phoneRaw: regData[i][regCols.phone],
          totalRemain: 0,
          memberships: []
        };
      }
      
      // 하나라도 진행중이면 활성 회원으로 정의
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
      
      // 잔여 횟수 합산
      memberInfoMap[phone].totalRemain += remain;
      
      if (membership) {
        memberInfoMap[phone].memberships.push(membership);
      }
    }
    
    // 2. 이미 문자함에 "대기" 중인 건이 있거나 최근 완료된 건 확인 (Standby 중복 생성 원천 가드 & 7일 재발송 주기 방어)
    var pendingSmsMap = {};
    var lastSentSmsDateMap = {};
    for (var j = 1; j < smsData.length; j++) {
      var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
      var sCategory = smsData[j][3];
      var sStatus = String(smsData[j][5]).trim();
      var sDateStr = smsData[j][0]; // "기록시간"
      
      if (sCategory === "복귀권유") {
        if (sStatus === "대기") {
          pendingSmsMap[sPhone] = true;
        } else if (sStatus === "완료") {
          var sDate = new Date(sDateStr);
          if (!lastSentSmsDateMap[sPhone] || sDate > lastSentSmsDateMap[sPhone]) {
            lastSentSmsDateMap[sPhone] = sDate;
          }
        }
      }
    }
    
    // 3. 미등록자 추출 및 세심한 분기 문자 생성
    var count = 0;
    var addedNames = [];
    
    var phones = Object.keys(memberInfoMap);
    for (var pIdx = 0; pIdx < phones.length; pIdx++) {
      var m = memberInfoMap[phones[pIdx]];
      
      // 활성(진행중) 회원이 아니고, 마지막 만료일이 있고 과거 시점이라면
      if (!m.isActive && m.lastExpire.getTime() > 0 && m.lastExpire < now) {
        
        // 문자함에 이미 복귀권유 "대기" 상태 문자가 들어있다면 절대 중복 생성 안 함!
        if (pendingSmsMap[phones[pIdx]]) continue;
        
        // 🚨 [스팸 방지 가드] 이미 복귀권유 문자가 최근 발송 완료된 이력이 있다면 7일간 재생성 차단
        var lastSentSmsDate = lastSentSmsDateMap[phones[pIdx]];
        if (lastSentSmsDate) {
          var diffSmsMs = now.getTime() - lastSentSmsDate.getTime();
          var daysSinceLastSms = diffSmsMs / (1000 * 60 * 60 * 24);
          if (daysSinceLastSms < 7) {
            continue; // 아직 발송 완료 후 7일이 경과하지 않았으므로 건너뜀
          }
        }
        
        // 만료 경과 일수 계산
        var diffMs = now.getTime() - m.lastExpire.getTime();
        var elapsedDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
        
        // 최소 1일 이상 지난 건들만 스캔 대상으로 삼음 (당일 만료 회원에 대한 당일 문자 폭격 방지)
        if (elapsedDays < 1) continue;
        
        var cleanName = m.name.replace(/\d{4}$/, ""); // 이름 끝 번호 제거
        var totalRemain = m.totalRemain;
        var formattedPhone = formatPhoneForSms(m.phoneRaw);
        
        // 🚨 [회원 이용 성향 스마트 분석 알고리즘 - v64.60 공용 헬퍼 적용]
        var mTypes = m.memberships || [];
        var attStats = memberAttendanceTypeMap[phones[pIdx]] || { therapy: 0, jumping: 0 };
        var regClass = determineMemberClassInfo(mTypes, attStats);
        var hasNoAttendance = (!memberAttendanceTypeMap[phones[pIdx]] || (attStats.therapy === 0 && attStats.jumping === 0));
        
        // 🚨 [원장님 스마트 가드] 4~6일차 만료 회원 하이브리드 발송 분기
        var isAppActive = hasMemberActiveInApp(phones[pIdx], 3); // 최근 3일 내 앱 활동 여부
        var targetChannel = "sms"; // 기본값 SMS
        
        if (elapsedDays >= 4 && elapsedDays <= 6) {
          if (isAppActive) {
            targetChannel = "noti"; // 앱 활성자는 무료 쪽지로 부드럽게 케어!
          } else {
            targetChannel = "sms"; // 앱 미활성자는 SMS로 강력 노킹!
          }
        }
        
        // AI 본문 동적 생성 (과거 히스토리 + 회원 이용 성향 맥락 인지형 + 만료일 유실 버그 원천 차단!)
        var formattedExpireDate = (m.lastExpire && m.lastExpire.getTime() > 0) ? Utilities.formatDate(m.lastExpire, "GMT+9", "yyyy-MM-dd") : "기록없음 (미입력)";
        var msg = generateWellnessAiSms(cleanName, totalRemain, elapsedDays, "복귀권유", formattedPhone, targetChannel, regClass, hasNoAttendance, formattedExpireDate);
        
        if (targetChannel === "noti") {
          // 무료 개인 쪽지 즉시 발송
          var notiTitle = "회원님, 남은 횟수(" + totalRemain + "회)를 지켜드릴게요! ❤️";
          if (totalRemain === 0) {
            notiTitle = "회원님, 함께 다져온 건강한 루틴을 계속 이어요! 😊";
          }
          sendPersonalNotification(formattedPhone, "복귀권유", notiTitle, msg);
        } else {
          // 오프라인 문자 대기열 적재
          smsSheet.appendRow([
            Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm"),
            cleanName,
            formattedPhone,
            "복귀권유",
            msg,
            "대기"
          ]);
        }
        
        count++;
        addedNames.push(cleanName);
      }
    }
    
    return { 
      success: true, 
      count: count, 
      message: count > 0 ? count + "명의 복귀 권유 대상을 세심하게 분류하여 문자 대기열에 쌓았습니다." : "새로운 복귀 권유 대상이 없습니다.",
      addedNames: addedNames
    };
    
  } catch (e) {
    return { error: "복귀 대상 추출 중 오류: " + e.toString() };
  }
}

/**
 * 📅 [v66.0] 관리자용 진행중인 정산 대상 회원 목록 조회 API (초고속 리턴)
 */
function getActiveMembersForSettle() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    if (!regSheet) return { error: "등록 현황 시트가 누락되었습니다." };
    
    var regData = regSheet.getDataRange().getDisplayValues();
    var regCols = getRegColumnIndices(regSheet);
    
    var activeMembers = [];
    var activePhonesMap = {}; 
    for (var i = 1; i < regData.length; i++) {
      var status = String(regData[i][regCols.status] || "").trim();
      if (status === "진행중" || status === "진행 중") {
        var phone = String(regData[i][regCols.phone] || "").replace(/[^0-9]/g, "");
        var name = regData[i][regCols.name];
        var membership = regData[i][regCols.membership] || "이용권";
        var remain = regData[i][regCols.remain] || "0";
        var expire = regData[i][regCols.expire] || "미정";
        
        if (!phone) continue;
        
        if (!activePhonesMap[phone]) {
          activePhonesMap[phone] = true;
          activeMembers.push({
            name: name,
            phone: phone,
            phoneRaw: regData[i][regCols.phone],
            membership: membership,
            remain: remain,
            expire: expire
          });
        }
      }
    }
    return { success: true, members: activeMembers };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * 📅 [v66.0] 개별 회원 1인에 대해 출석기록 분석, 이전 이력 대조 및 AI 평가를 수행하는 청크 생성 API (실행시간 2초, 타임아웃 0%)
 */
function generateIndividualAttendanceSms(member, startDateStr, endDateStr, baseDateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    var smsSheet = ss.getSheetByName("문자발송");
    
    if (!logSheet || !smsSheet) {
      return { error: "필요한 시트(출석기록/문자발송)가 누락되었습니다." };
    }
    
    var logData = logSheet.getDataRange().getDisplayValues();
    var smsData = smsSheet.getDataRange().getDisplayValues();
    var logCols = getAttendanceColumnIndices(logSheet);
    
    var startD = new Date(startDateStr);
    var endD = new Date(endDateStr);
    startD.setHours(0,0,0,0);
    endD.setHours(23,59,59,999);
    
    var phone = member.phone;
    var now = new Date();
    
    // A. 중복 가드 및 이전 정산 이력 단 1건만 역순 스캔 추출
    var isDuplicate = false;
    var prevSms = "";
    for (var j = 1; j < smsData.length; j++) {
      var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
      var sCategory = String(smsData[j][3] || "").trim();
      var sStatus = String(smsData[j][5] || "").trim();
      var sContent = String(smsData[j][4] || "");
      var sTime = String(smsData[j][0] || "");
      
      if (sPhone === phone && sCategory === "출석정산") {
        if (sStatus === "대기" && sTime.includes(Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd"))) {
          isDuplicate = true;
        }
        prevSms = sContent;
      }
    }
    
    if (isDuplicate) return { success: true, status: "duplicate" };
    
    // B. 출석기록 시트에서 지정 기간 내 출석 집계
    var stats = { total: 0, jumping: 0, therapy: 0 };
    for (var k = 1; k < logData.length; k++) {
      var lPhone = String(logData[k][logCols.phone] || "").replace(/[^0-9]/g, "");
      var lDateStr = logData[k][logCols.date];
      var lType = String(logData[k][logCols.type] || ""); 
      var lReason = String(logData[k][logCols.reason] || ""); 
      
      if (lPhone === phone && lDateStr) {
        var lDate = new Date(lDateStr);
        if (lDate >= startD && lDate <= endD) {
          stats.total++;
          
          var isCombo = (lType === "복합" || lType.indexOf("복합") !== -1 || lReason.indexOf("복합") !== -1);
          var isTherapy = (lType.indexOf("테라피") !== -1 || lType.indexOf("원적외선") !== -1 || lType.indexOf("반신욕") !== -1 || lType.indexOf("보너스") !== -1 || lReason.indexOf("테라피") !== -1 || lReason.indexOf("보너스") !== -1) && !isCombo;
          
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
    }
    
    // C. Gemini AI 평가 코멘트 생성
    var cleanName = member.name.replace(/\d{4}$/, "");
    var aiComment = generateAttendanceReportAiComment(
      cleanName, 
      stats, 
      prevSms, 
      member.remain, 
      member.expire, 
      startDateStr, 
      endDateStr
    );
    
    // D. 최종 문자 조립 (정산 기준 시각 HH:mm 현재 포맷 주입으로 명확성 극대화)
    var smsBody = "[노형점핑 웰니스 정산]\n" +
                  "📅 정산 기간: " + startDateStr + " ~ " + endDateStr + "\n" +
                  "🏃 출석 기록: 총 " + stats.total + "회 (점핑 " + stats.jumping + "회 / 테라피 " + stats.therapy + "회)\n\n" +
                  "💌 웰니스 코칭 레터:\n" +
                  aiComment + "\n\n" +
                  "🎫 회원권 현황 (정산 기준 시점: " + baseDateStr + " " + Utilities.formatDate(now, "GMT+9", "HH:mm") + " 현재):\n" +
                  "- 잔여 횟수: " + member.remain + "회\n" +
                  "- 이용 만료일: " + member.expire;
    
    var formattedPhone = formatPhoneForSms(member.phoneRaw);
    smsSheet.appendRow([
      Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm"),
      cleanName,
      formattedPhone,
      "출석정산", 
      smsBody,
      "대기"
    ]);
    
    return { success: true, status: "created", name: cleanName };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * 🤖 [v66.0] 제미나이를 통한 개별 회원 맞춤형 한 달 출석 평가 코멘트 생성기
 */
function generateAttendanceReportAiComment(name, stats, prevSms, remain, expire, startDateStr, endDateStr) {
  try {
    // 정산 시작일에서 실제 정산 대상 월 추출 (예: "2026-05-01" -> 5월)
    var startD = new Date(startDateStr);
    var targetMonthNum = startD.getMonth() + 1;
    var nextMonthNum = targetMonthNum === 12 ? 1 : targetMonthNum + 1;
    var targetMonthStr = targetMonthNum + "월";
    var nextMonthStr = nextMonthNum + "월";

    var systemInstruction = "당신은 제주 노형점핑클럽의 따뜻하고 쾌활하며 전문적인 웰니스 코치입니다. " +
                            "회원의 " + targetMonthStr + " 한 달 동안의 출석 기록과 상태를 보고, 첫 인사부터 " + nextMonthStr + " 마무리 응원까지 매끄럽게 이어지는 한 편의 따뜻한 맞춤형 코칭 편지(줄글 형태)를 직접 작성하십시오. " +
                            "답변은 반드시 한국어 경어체로 180자~300자 내외로 자연스럽고 풍부하게 작성해 주시고, " +
                            "정산하는 시점(오늘)이 몇 월 며칠이든 상관없이, 회원이 돌아보는 대상은 반드시 '" + targetMonthStr + "'이고 다짐을 하는 다가오는 달은 '" + nextMonthStr + "'이어야 하므로, 날짜 혼선 없이 반드시 편지 내에서 이번 달은 '" + targetMonthStr + "', 다음 달은 '" + nextMonthStr + "'로 명시하여 대화하십시오. " +
                            "문장 곳곳에 하트(❤️), 응원(💪), 미소(🥰)와 같은 따뜻한 이모지를 2~3개 적절히 섞어 코치의 진심어린 온기를 더해 주십시오.";
    
    var prompt = "■ 회원 및 출석 정보\n" +
                 "- 회원 이름: " + name + "\n" +
                 "- 정산 대상 월 (이번 달): " + targetMonthStr + "\n" +
                 "- 다음 활동 월 (다음 달): " + nextMonthStr + "\n" +
                 "- 정산 기간: " + startDateStr + " ~ " + endDateStr + "\n" +
                 "- 이번 출석 횟수: 총 " + stats.total + "회 (점핑 운동: " + stats.jumping + "회, 원적외선 테라피: " + stats.therapy + "회)\n" +
                 "- 회원권 남은 횟수: " + remain + "회\n" +
                 "- 회원권 이용 마감일: " + expire + "\n\n" +
                 "■ 지난 정산 문자 기록 (가장 최근 발송된 히스토리)\n" +
                 (prevSms ? prevSms : "(과거 출석정산 이력 없음 - 첫 정산 대상)") + "\n\n" +
                 "■ 핵심 지시사항 및 편지 구성 요소 (필수 준수)\n" +
                 "1. 편지 구조 정의:\n" +
                 "   - [첫인사]: 회원님의 이름을 다정하고 따뜻하게 부르며 시작하는 친근한 첫인사. 오늘이 몇 월 며칠이든 관계없이 편지 내용 상에서는 반드시 " + targetMonthStr + "의 건강 여정 정산 소식임을 명시하십시오. (예: '♥지은 회원님, 따뜻하고 눈부셨던 " + targetMonthStr + "도 건강하게 잘 마감하셨나요? 🌟 한 달간의 소중한 웰니스 성적표를 들고 코치가 찾아왔습니다!')\n" +
                 "   - [본문 (출석 분석)]: " + targetMonthStr + " 출석 빈도와 남은 횟수, 그리고 이전 히스토리 대비 증감을 분석하여 칭찬/공감/우려를 전문적이면서 다정하게 언급 (예: 지난달에 비해 자주 오셔서 정말 기쁘다든지, 혹은 바쁘신지 출석이 뜸해 소중한 회원권 소멸 걱정이 된다든지)\n" +
                 "   - [끝인사]: 다가오는 " + nextMonthStr + "에도 힘차게 건강 관리를 이어나갈 수 있도록 따뜻하게 격려하고 클럽에서 반갑게 만나자는 다정한 끝인사. (예: '바쁘신 일상 속에서도 자신을 위한 소중한 시간을 매일 개척해 나가시는 모습이 진정으로 멋집니다. 다가오는 " + nextMonthStr + "에도 늘 정성을 다하는 코치로 대기하고 있을게요! 클럽에서 뵙겠습니다. ❤️')\n" +
                 "2. 출석 빈도별 반응 룰:\n" +
                 "   - 출석 0~4회: 바쁜 일상에 깊이 공감해주되, 소중한 남은 횟수(" + remain + "회)와 마감일(" + expire + ")을 강조하며 이용 기간 내에 다 쓰지 못해 소멸할까 걱정되는 코치의 진심 어린 염려 톤 부각\n" +
                 "   - 출석 5~12회: 점핑과 테라피를 적절히 조화롭게 활용하는 습관을 칭찬하고 꾸준한 건강 루틴화 독려\n" +
                 "   - 출석 13회 이상: 엄청난 성실함과 뜨거운 건강 열정에 감동받은 코치의 극찬과 경탄의 찬사\n" +
                 "3. 지난 정산 이력과의 1:1 대조 분석:\n" +
                 "   - 지난 문자 기록이 있다면 꼭 이전 출석 횟수와 비교하여 증가/감소의 변화 양상을 극찬하거나 공감하며 안부 묻기\n" +
                 "4. 출력 형식:\n" +
                 "   - 정산 기간이나 출석 기록 테이블, 회원권 남은 횟수 테이블 등 표 형태의 정보는 시스템이 앞뒤에 깔끔하게 부착하므로, 이곳에서는 **오직 회원의 감성을 깊이 어루만지는 완성도 높은 줄글 형태의 '편지 본문 전체(첫인사 + 분석 + 끝인사)'만 180~300자 내외로 깔끔하게 작성**하여 반환하십시오. 절대 정산하는 날짜(6월 등) 기준의 헷갈리는 문장을 적지 말고, 명확히 " + targetMonthStr + "과 " + nextMonthStr + "으로 고정하여 작성해 주십시오.";
    
    var aiComment = callGeminiBackend(prompt, systemInstruction);
    
    if (aiComment) {
      aiComment = aiComment.trim().replace(/^["']|["']$/g, "");
      return aiComment;
    }
  } catch (e) {
    Logger.log("Gemini 정산 코멘트 생성 에러: " + e.toString());
  }
  
  var total = stats.total;
  var jumping = stats.jumping;
  var therapy = stats.therapy;
  var comment = "";
  
  // 1. 출석 성향 설명 동적 빌드 (점핑만 온 사람, 테라피만 온 사람, 둘 다 온 사람 분기)
  var habitDesc = "건강을 세심히 가꾸시는 모습이 정말 멋지십니다!";
  if (jumping > 0 && therapy === 0) {
    habitDesc = "신나는 음악 소리에 몸을 실어 트램폴린 점핑 운동에 온전히 집중하신 열정적인 건강 루틴이 참으로 인상적입니다!";
  } else if (jumping === 0 && therapy > 0) {
    habitDesc = "아늑한 편백실 원적외선 테라피실에서 개운하게 땀을 빼며 지친 몸과 마음의 피로를 세심히 다스리신 힐링 루틴이 참으로 아름답습니다!";
  } else if (jumping > 0 && therapy > 0) {
    habitDesc = "땀을 쫙 흘리는 점핑 운동의 강력한 활력과 온몸을 사르르 녹여주는 테라피의 편안한 피로 회복을 지혜롭게 조화시키신 멋진 하이브리드 건강 루틴을 실천하셨습니다!";
  }
  
  if (total === 0) {
    comment = name + " 회원님, 지난 정산 기간 동안 클럽에서 한 번도 뵙지 못해 코치가 무척 안타깝고 염려스러운 마음이 가득한 한 달이었습니다. 😢 바쁘신 일상이 정리되는 대로, 소중한 남은 회원권이 만료일 전에 아깝게 소멸되지 않도록 꼭 다시 클럽에 들러 건강 충전을 시작해 보세요!";
  } else if (total <= 5) {
    comment = name + " 회원님, 이번 정산 기간 동안 총 " + total + "회의 출석을 기록해 주셨네요! 바쁜 일정 속에서도 걸음해 주셔서 감사하지만, 남은 회원권 잔여 횟수가 " + remain + "회나 유효하게 남아있어 이용 만료일(" + expire + ")까지 모두 소진하시지 못해 아깝게 소멸될까 코치가 진심으로 걱정하고 있습니다. 다음 달은 조금만 더 자주 뵈어요! ❤️";
  } else if (total <= 12) {
    comment = name + " 회원님, 지난 정산 기간 동안 총 " + total + "회의 소중한 출석을 달성하셨네요! " + habitDesc + " 회원님만의 건강한 리듬이 삶 속에 점차 단단하게 자리를 잡아가고 있는 것 같아 코치도 든든하고 뿌듯한 보람을 느낍니다. 늘 응원하고 자랑스럽습니다. 😊";
  } else {
    comment = name + " 회원님, 이번 정산 기간 동안 무려 " + total + "회라는 경이롭고 뜨거운 출석 대기록을 돌파하셨습니다! 대단하십니다. 🏆💪 " + habitDesc + " 회원님의 흔들림 없는 꾸준함과 뜨거운 건강 집념에 코치도 엄청난 감동과 웰니스 에너지를 선물 받았습니다. 삶의 진정한 성취를 보여주시는 회원님을 온 맘 다해 존경하고 축복합니다!";
  }
  
  return comment;
}

/**
 * [관리자 전용] 4~6일 연속 결석으로 점수가 차감 중인 회원 추출 및 문자 생성 (v46.35)
 */
function checkInactivityDebuffAbsentees() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    var logSheet = ss.getSheetByName("출석기록");
    var smsSheet = ss.getSheetByName("문자발송");
    
    if (!regSheet || !logSheet || !smsSheet) return { error: "필요한 시트(등록현황/출석기록/문자발송)가 없습니다." };
    
    // [중복 업데이트 및 실시간 정밀 타임라인 덮어쓰기 알고리즘]
    // 1. 기존에 시트에 존재하는 모든 '대기' 상태 of '출석디버프' 문자열 행들을 스캔하여 완벽히 삭제(Clean-up)합니다!
    var smsDataForCleanup = smsSheet.getDataRange().getValues();
    for (var i = smsDataForCleanup.length - 1; i >= 1; i--) {
      var category = String(smsDataForCleanup[i][3]);
      var status = String(smsDataForCleanup[i][5]).trim();
      if (category === "출석디버프" && status === "대기") {
        smsSheet.deleteRow(i + 1);
      }
    }
    
    // 삭제 처리 완료 후 최신 데이터 로드!
    var regData = regSheet.getDataRange().getDisplayValues();
    var logData = logSheet.getDataRange().getValues();
    var smsData = smsSheet.getDataRange().getDisplayValues();
    
    var regCols = getRegColumnIndices(regSheet);
    var logCols = getAttendanceColumnIndices(logSheet);
    
    var now = new Date();
    
    // 1. 마지막 출석일 맵 작성 (폰번호 기준)
    var lastAttendanceMap = {};
    for (var i = 1; i < logData.length; i++) {
      var phone = String(logData[i][logCols.phone] || "").replace(/[^0-9]/g, "");
      var dRaw = logData[i][logCols.date];
      if (!phone || !dRaw) continue;
      
      var attendanceDate = (dRaw instanceof Date) ? dRaw : new Date(dRaw);
      if (!lastAttendanceMap[phone] || attendanceDate > lastAttendanceMap[phone]) {
        lastAttendanceMap[phone] = attendanceDate;
      }
    }
    
    // 2. 이미 발송 대기 중이거나 완료된 목록 확인 (중복 생성 방지 및 7일 주기 재발송용)
    var hasPendingSmsMap = {};
    var lastSentSmsDateMap = {};
    for (var j = 1; j < smsData.length; j++) {
      var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
      var sCategory = smsData[j][3];
      var sStatus = smsData[j][5];
      var sDateStr = smsData[j][0]; // "기록시간"
      
      if (sCategory === "출석디버프") {
        if (sStatus === "대기") {
          hasPendingSmsMap[sPhone] = true;
        } else if (sStatus === "완료") {
          var sDate = new Date(sDateStr);
          if (!lastSentSmsDateMap[sPhone] || sDate > lastSentSmsDateMap[sPhone]) {
            lastSentSmsDateMap[sPhone] = sDate;
          }
        }
      }
    }
    
    // 3. 연속 결석 4~6일인 회원 추출 및 문자 생성
    var count = 0;
    var addedNames = [];
    
    var midnightNow = new Date();
    midnightNow.setHours(0,0,0,0);
    
    for (var k = 1; k < regData.length; k++) {
      var status = String(regData[k][regCols.status]).trim();
      if (status !== "진행중" && status !== "진행 중") continue;
      
      var phone = String(regData[k][regCols.phone] || "").replace(/[^0-9]/g, "");
      var name = regData[k][regCols.name];
      if (!phone) continue;
      
      var lastDate = lastAttendanceMap[phone];
      if (!lastDate) continue; // 출석 기록이 전혀 없는 신규 회원은 대상 제외
      
      var midnightLast = new Date(lastDate.getTime());
      midnightLast.setHours(0,0,0,0);
      
      var inactiveDays = calculateInactiveDays(midnightLast, midnightNow);
      
      // 4~6일 결석일 때 문자 생성 (디버프 경고 전송 적기)
      if (inactiveDays >= 4 && inactiveDays <= 6) {
        if (hasPendingSmsMap[phone]) continue;
        
        var lastSentSmsDate = lastSentSmsDateMap[phone];
        if (lastSentSmsDate) {
          if (lastDate && lastDate <= lastSentSmsDate) {
            var diffMs = now.getTime() - lastSentSmsDate.getTime();
            var daysSinceLastSms = diffMs / (1000 * 60 * 60 * 24);
            if (daysSinceLastSms < 7) {
              continue; // 아직 7일이 지나지 않았으므로 건너뜀
            }
          }
        }
        
        var cleanName = name.replace(/\d{4}$/, ""); // 이름 뒤 번호 제거
        var penaltyVal = (inactiveDays - 3) * 100;
        
        var notiTitle = "🚨 웰니스 점수 방전 경고! 홈 케어 미션 🔥";
        var notiContent = cleanName + " 회원님! 운동을 쉬신 지 벌써 " + inactiveDays + "일이 지나, 아쉽게도 웰니스 누적 점수가 매일 100 EXP씩 방전(감점)되기 시작했어요! 😭 누적 -" + penaltyVal + " EXP 상태입니다.\n\n하지만 걱정 마세요! 오늘 클럽에 오셔서 신나게 점핑 뛰고 출석체크만 쾅! 하시면 깎였던 모든 점수와 순위가 즉시 100% 마법처럼 전부 복구(부활)됩니다! ✨\n\n만약 방문이 어려우시다면 홈 케어로 가벼운 스트레칭이나 물 2L 마시기 미션이라도 실천해보시는 건 어떨까요? 오늘 꼭 앱에서 일상 기록을 남겨 에너지를 지키세요! ❤️";
        
        var isActive = isAppActiveDuringAbsence(phone, lastDate, now);
        
        // A. 인앱 쪽지 발송 (오늘 이미 안 보낸 경우)
        try {
          if (!hasSentNotificationToday(phone, "디버프", "웰니스 점수 방전 경고")) {
            sendPersonalNotification(phone, "디버프", notiTitle, notiContent);
          }
        } catch (notiErr) {
          Logger.log("디버프 인앱 쪽지 생성 실패 (" + cleanName + "): " + notiErr.toString());
        }
        
        // B. 앱 활성 상태인 경우 SMS 발송 스킵
        if (isActive) {
          continue; 
        }
        
        var msg = "🚨 [노형점핑] " + cleanName + " 회원님! 운동을 쉬신 지 벌써 " + inactiveDays + "일이 지나, 아쉽게도 웰니스 누적 점수가 매일 100 EXP씩 방전(감점)되기 시작했어요! 😭 누적 -" + penaltyVal + " EXP 상태입니다. 하지만 걱정 마세요! 오늘 클럽에 오셔서 신나게 점핑 뛰고 출석체크만 쾅! 하시면 깎였던 모든 점수와 순위가 즉시 100% 마법처럼 전부 복구(부활)됩니다! ✨ 오늘 꼭 오셔서 건강 충전해 가세요! ❤️";
        
        var formattedPhone = formatPhoneForSms(regData[k][regCols.phone]);
        
        smsSheet.appendRow([
          Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm"),
          cleanName,
          formattedPhone,
          "출석디버프",
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
      message: count > 0 ? count + "명의 디버프 경고 대상을 추출했습니다." : "새로운 디버프 경고 대상이 없습니다.",
      addedNames: addedNames
    };
  } catch (e) {
    return { error: "디버프 추출 오류: " + e.toString() };
  }
}

// [이동 완료] normalizePhoneDigits, formatPhoneNumber, parseWeightSafely 함수들은 Common_Utils.gs로 이동되었습니다.

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

// [이동 완료] formatPhoneForSms 함수는 Common_Utils.gs로 이동되었습니다.

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

// [v58.5] 중복 정의된 getPillarNotice 및 updatePillarNotice 함수는 백엔드 무결성을 위해 하단의 최신 공용 함수군으로 통합 단일화되었습니다.

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
// [이동 완료] recordAppAccess 및 getDailyActiveAdventurers 함수는 Bridge_Login.gs로 이동되었습니다.

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

/**
 * ⚡ 클럽 출석 기록 동기화 (v44.185)
 * [출석기록] 시트의 K열(운동타임수)을 분석하여 보상을 지급합니다.
 */
function syncClubRecord(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    var actSheet = ss.getSheetByName("활동기록") || ss.insertSheet("활동기록");
    
    if (!logSheet) return { success: false, error: "'출석기록' 시트가 없습니다." };
    
    var phone = normalizePhoneDigits(payload.phone);
    var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    
    // 1. 이미 오늘 동기화했는지 확인
    var actData = actSheet.getDataRange().getDisplayValues();
    for (var i = 1; i < actData.length; i++) {
      var aDateRaw = actData[i][0];
      var aPhone = normalizePhoneDigits(actData[i][3]);
      var aType = String(actData[i][4]);
      
      // 날짜 비교 (하이픈 처리 등 정규화)
      var dMatch = aDateRaw.match(/(\d{4})[^\d]+(\d{1,2})[^\d]+(\d{1,2})/);
      var normADate = dMatch ? dMatch[1] + "-" + dMatch[2].padStart(2, '0') + "-" + dMatch[3].padStart(2, '0') : "";
      
      if (normADate === todayStr && aPhone === phone && aType === "출석동기화") {
        return { success: false, error: "이미 오늘자 출석 동기화가 완료되었습니다." };
      }
    }
    
    // 2. 출석기록 시트에서 오늘 기록 찾기
    var logData = logSheet.getDataRange().getDisplayValues();
    var todayRecord = null;
    for (var j = logData.length - 1; j >= 1; j--) {
      var lDateMatch = logData[j][0].match(/(\d{4})[^\d]+(\d{1,2})[^\d]+(\d{1,2})/);
      var lDateStr = lDateMatch ? lDateMatch[1] + "-" + lDateMatch[2].padStart(2, '0') + "-" + lDateMatch[3].padStart(2, '0') : "";
      var lPhone = normalizePhoneDigits(logData[j][3]);
      
      if (lDateStr === todayStr && lPhone === phone) {
        todayRecord = logData[j];
        break; // 최신 기록 하나만
      }
    }
    
    if (!todayRecord) return { success: false, error: "오늘 클럽 출석 기록이 확인되지 않습니다. 키오스크 출석을 먼저 해주세요." };
    
    // 3. 점수 계산 (소수점 타임 비례 계산 및 복합출석 더블 보상 연동)
    var visitPoints = 20;
    var timeCount = parseFloat(todayRecord[10] || 0); // K열: 운동타임수
    var reasonText = String(todayRecord[8] || ""); // I열: 사유
    var isCombo = (reasonText.indexOf("복합") !== -1);
    
    var timePoints = Math.round(timeCount * 40);
    var therapyPoints = (timeCount === 0 || isCombo) ? 30 : 0;
    var totalPoints = visitPoints + timePoints + therapyPoints;
    var memberName = todayRecord[2];
    
    var detailMsg = "운동 타임: " + timeCount;
    if (isCombo) {
      detailMsg += " (복합출석: 점핑 + 테라피)";
    } else if (timeCount === 0) {
      detailMsg = "테라피 완료";
    }
    
    // 4. 활동기록에 저장
    var now = new Date();
    actSheet.appendRow([
      now,
      Utilities.formatDate(now, "GMT+9", "HH:mm:ss"),
      memberName,
      "'" + phone,
      "출석동기화",
      "클럽 방문",
      detailMsg,
      "",
      totalPoints
    ]);
    
    return { success: true, points: totalPoints, timePoints: timePoints, therapyPoints: therapyPoints };
    
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * [v45.7] 일일 활동 기록 시스템 (정밀 머리말 태깅)
 */
function recordActivityLog(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("일일_활동_기록") || ss.insertSheet("일일_활동_기록");
    
    if (sheet.getLastRow() === 0) {
      // [v45.8] 10개 컬럼 최적화 (원장님 기획안)
      var headers = ["날짜", "전화번호", "이름", "센터방문_실천", "운동강도_실천", "일반실천_합산", "일반회복_합산", "체력보너스_합산", "완료내역", "웰니스총점"];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f1f3f5");
    }

    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var phone = normalizePhoneDigits(payload.phone);
    var type = payload.type || "일반"; 
    var score = Number(payload.score || 0);

    // [v45.8] 머리말 및 액션 결정 (v45.9 정밀 매핑)
    var statTag = "[실천]"; // 기본은 실천으로 변경 (체력 보너스 남용 방지)
    
    // 1순위: 명시적인 statType 우선
    if (payload.statType === "perf") statTag = "[실천]";
    else if (payload.statType === "def") statTag = "[회복]";
    else if (payload.statType === "health") statTag = "[체력]";
    else {
      // 2순위: 타입 및 항목명으로 추론
      if (type === "습관" || type === "오아시스" || type === "방어" || type === "회복" || type === "로그인" || payload.item === "셀프 칭찬" || payload.item.indexOf("인증") > -1) {
        // [v45.9] 아카이브/오아시스/보너스/로그인은 명시적 타입에 따라 '회복'(회복력)으로 취급
        statTag = "[회복]";
      } else if (type === "퀘스트" || type === "식단" || type === "출석" || type === "수행" || type === "실천") {
        statTag = "[실천]";
      } else if (type === "보너스") {
        statTag = "[체력]";
      }
    }

    var action = payload.action || (type === "습관" ? "체크" : "인증");
    var detailInfo = statTag + " " + payload.item + " " + action + "(" + score + ")";

    if (type === "로그인") {
      action = "";
      detailInfo = "[회복] 로그인 체크(5)";
    }

    // 1. 행 찾기 및 중복 체크
    var data = sheet.getDataRange().getValues();
    var targetRowIdx = -1;
    var targetRowIdxInArray = -1;
    var completedDetails = "";
    for (var i = 1; i < data.length; i++) {
      var rowDate = (data[i][0] instanceof Date) ? Utilities.formatDate(data[i][0], "GMT+9", "yyyy-MM-dd") : String(data[i][0]);
      if (rowDate === todayStr && normalizePhoneDigits(data[i][1]) === phone) {
        targetRowIdx = i + 1;
        targetRowIdxInArray = i;
        completedDetails = String(data[i][8] || ""); // I열(9번째)
        break;
      }
    }

    // [v45.8] 1일 1회 원칙 (보너스 포함 중복 방지)
    if (completedDetails.indexOf(payload.item + " " + action) > -1) {
      return { success: true, message: "이미 기록됨" };
    }

    // 2. 기록 (메모리 최적화 일괄 적용)
    if (targetRowIdx > -1) {
      // 기존 행 업데이트 (메모리 버퍼에서 직접 연산 후 단 1회 setValues)
      var rowValues = data[targetRowIdxInArray];
      
      // 혹시라도 array 길이가 부족한 경우 방지
      while (rowValues.length < 10) {
        rowValues.push(0);
      }
      
      // [자가 치유] 기존 행에 혹시라도 0이 잘린 전화번호가 들어가 있다면 올바른 텍스트로 보정 및 업데이트 강제
      rowValues[1] = "'" + normalizePhoneDigits(rowValues[1]);
      
      if (type === "로그인") {
        rowValues[6] = Number(rowValues[6] || 0) + 5; // 일반회복_합산
      } else {
        var colIdxInArray = 5; // 일반실천_합산 기본 (0-indexed)
        if (type === "출석") {
          if (payload.item === "센터방문") colIdxInArray = 3;
          else if (payload.item === "운동강도") colIdxInArray = 4;
        } else if (statTag === "[회복]") {
          colIdxInArray = 6;
        } else if (statTag === "[실천]") {
          colIdxInArray = 5;
        } else if (statTag === "[체력]") {
          colIdxInArray = 7;
        }
        rowValues[colIdxInArray] = Number(rowValues[colIdxInArray] || 0) + score;
      }
      
      // 완료내역 및 웰니스총점 업데이트
      var currentDetails = String(rowValues[8] || "");
      if (currentDetails.indexOf(detailInfo) === -1) {
        rowValues[8] = currentDetails ? currentDetails + ", " + detailInfo : detailInfo;
      }
      
      // D(3) ~ H(7) 열 합산
      rowValues[9] = Number(rowValues[3] || 0) + Number(rowValues[4] || 0) + Number(rowValues[5] || 0) + Number(rowValues[6] || 0) + Number(rowValues[7] || 0);
      
      // 단 1회의 Sheets 쓰기 호출 (기존 6회에서 1회로 단축!)
      sheet.getRange(targetRowIdx, 1, 1, 10).setValues([rowValues]);
    } else {
      // 새 행 추가 (기존 1회 call 유지)
      var newRow = [todayStr, "'" + phone, payload.name, 0, 0, 0, 0, 0, "", 0];
      if (type === "로그인") {
        newRow[6] = 5;
        newRow[8] = "[회복] 로그인 체크(5)";
        newRow[9] = 5;
      } else {
        var colIdxInArray = 5;
        if (type === "출석") {
          if (payload.item === "센터방문") colIdxInArray = 3;
          else if (payload.item === "운동강도") colIdxInArray = 4;
        } else if (statTag === "[회복]") {
          colIdxInArray = 6;
        } else if (statTag === "[실천]") {
          colIdxInArray = 5;
        } else if (statTag === "[체력]") {
          colIdxInArray = 7;
        }
        newRow[colIdxInArray] = score;
        newRow[8] = detailInfo;
        newRow[9] = score;
      }
      sheet.appendRow(newRow);
    }
    
    clearUserDashboardCache(phone);
    return { success: true };
  } catch (e) { return { success: false, error: e.toString() }; }
}

/**
 * [v45.9] 마스터 플랜 기반 인바디 기록 저장 및 점수 계산
 */
function DEPRECATED_submitInBodyRecord(payload) {
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
    
    var score = 0;
    if (prevRecord) {
      var wDiff = prevRecord.weight - weight; // 감량 시 양수
      var mDiff = muscle - prevRecord.muscle; // 증량 시 양수
      var fDiff = prevRecord.fat - fat;       // 감량 시 양수
      
      // [v45.9] 마스터 플랜 배점 (수정: 체지방 1%당 1000점)
      // 1. 체중: -100g당 +50
      if (wDiff !== 0) score += Math.floor(wDiff * 10) * 50; 
      // 2. 골격근량: +100g당 +200 (1kg당 2000점)
      if (mDiff !== 0) score += Math.floor(mDiff * 10) * 200;
      // 3. 체지방률: -0.1%당 +100 (1%당 1000점 - 체중의 2배 가치)
      if (fDiff !== 0) score += Math.floor(fDiff * 10) * 100;
      
      // 4. 유지 보너스: 전주 대비 ±0.2kg 내 변화 시 +100
      if (Math.abs(wDiff) <= 0.2) score += 100;
    }
    
    sheet.appendRow([new Date(), name, "'" + phone, weight, muscle, fat, score]);
    

    
    return { success: true, score: score };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// [v46.20] 돌발퀘스트 & 글리코겐 방패 시스템 마스터 엔진
// ==========================================

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

// [perf] ss 객체, logData, memberName을 외부에서 공유받아 SpreadsheetApp 및 시트 읽기 중복 방지
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
 * ==========================================
 * [v46.20] 마을 이장의 집 (Chief's Sanctum) 관리 엔진
 * ==========================================
 */

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

// [이동 완료] getDailyActiveAdventurers duplicate 함수는 Bridge_Login.gs로 단일화되어 이동되었습니다.

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

/**
 * 5. 축복 대기 중인 최근 인증 내역 조회
 */
function getRecentCertifications() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("아카이브");
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // A열부터 K열(11열)까지 데이터 조회
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getDisplayValues();
    var results = [];
    
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var type = String(row[4] || "");
      var blessStatus = String(row[10] || ""); // K열: 축복상태
      
      // 시각 유형이고 축복 완료가 아닌 행만 필터링
      if ((type === "퀘스트" || type === "습관" || type === "식단") && blessStatus !== "축복완료") {
        results.push({
          rowIdx: i + 2,
          date: String(row[0] || ""),
          time: String(row[1] || ""),
          name: String(row[2] || ""),
          phone: String(row[3] || ""),
          type: type,
          item: String(row[5] || ""),
          content: String(row[6] || "")
        });
      }
    }
    // 최신 글이 위로 오도록 반전
    results.reverse();
    return results.slice(0, 20); // 성능 최적화를 위해 상위 20개만 반환
  } catch (e) {
    return [];
  }
}

/**
 * 6. 모험가 항목 개별 축복 (보너스 +5 EXP 및 K열 업데이트)
 */
function blessAction(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("아카이브");
    if (!sheet) return { success: false, error: "'아카이브' 시트를 찾을 수 없습니다." };
    
    var rowIdx = parseInt(payload.rowIdx);
    if (isNaN(rowIdx) || rowIdx < 2) return { success: false, error: "유효하지 않은 행 인덱스입니다." };
    
    // K열(11열)에 "축복완료" 상태 기록
    sheet.getRange(rowIdx, 11).setValue("축복완료");
    
    // +5 EXP 활동기록으로 즉시 부여 (방어 스탯에 반영)
    recordActivityLog({
      phone: payload.phone,
      name: payload.name,
      type: "방어",
      item: "✨ 이장의 축복 보너스",
      score: 5,
      action: "축복",
      statType: "def"
    });
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 7. 이장의 즉석 돌발 퀘스트 선포 및 전령의 기둥 연동
 */
function createSurpriseQuest(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = checkAndCreateQuestRegistrySheet();
    
    var now = new Date();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var durationMs = (parseInt(payload.durationMinutes) || 60) * 60 * 1000;
    var endTime = new Date(now.getTime() + durationMs);
    var endTimeStr = Utilities.formatDate(endTime, "GMT+9", "yyyy-MM-dd HH:mm:ss");
    
    var nextId = "SUR_" + Date.now();
    
    var qScore = payload.score ? Number(payload.score) : 15;
    var qMethod = payload.method || "사진";
    
    // 퀘스트 행 등록
    sheet.appendRow([
      nextId,
      todayStr,
      "이장",
      payload.title,
      payload.description || "마을 이장님의 특별 계시입니다.",
      endTimeStr,
      "", // G열: 전화번호 공란 (전체 멤버용 돌발)
      "진행중", // H열: 상태
      qScore,   // I열: 점수
      qMethod   // J열: 인증방식
    ]);
    
    // 전령의 기둥(실시간 공지)으로 연동 선포
    var noticeContent = "⚡ [돌발 계시] " + payload.title;
    if (payload.description) {
      noticeContent += "! " + payload.description;
    }
    updatePillarNotice({ content: noticeContent });
    
    // [v64.50] 즉석 돌발 퀘스트 생성 시에도 '전체_알림_목록' 시트에 글로벌 레코드로 1행 즉시 적재 발송합니다!
    var globalTitle = "⚡ [돌발 퀘스트] " + payload.title + " 선포! ✉️";
    var globalContent = "📢 [웰니스 코치의 긴급 돌발 퀘스트 선포!] 📢\n\n" +
                        "오늘의 돌발 퀘스트가 선포되었습니다! 오늘 자정까지 완수하고 추가 보너스 에너지를 획득하세요!\n\n" +
                        "🔥 [돌발 퀘스트]: " + payload.title + "\n" +
                        "📝 [임무 설명]: " + (payload.description || "웰니스 코치의 특별 계시입니다.") + "\n" +
                        "💎 [보상 EXP]: +" + qScore + " EXP\n" +
                        "🎯 [인증 방법]: " + qMethod + " 인증\n\n" +
                        "지금 즉시 대시보드에서 돌발 퀘스트를 인증하고 수호 점수를 획득해 보세요! ⚔️";
    sendGlobalNotification("quest", globalTitle, globalContent);
    
    return { success: true, endTime: endTime.getTime() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// [v46.32] 주간 다회 출석 체력 보너스 자동 누적 정산 엔진 (3회: 100, 4회: 200, 5회: 300, 6회: 500 EXP)
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

// [v65.00] 구글의 소스코드 노출 자동차단을 예방하는 안전한 백엔드 API 키 저장/배포 시스템 (복수 키 로테이션 풀 지원)
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
 * 📊 [v64.40] 특정 회원의 최근 발송 히스토리(문자발송 + 개인쪽지) 통합 수집기
 * 대상 전화번호의 최근 4건의 발송 내역을 가져와 AI 비서에게 맥락으로 전달합니다.
 */
function getMemberMessageHistory(phone) {
  try {
    var cleanPhone = String(phone || "").replace(/[^0-9]/g, "");
    if (!cleanPhone) return "최근 발송 이력 없음";
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var smsSheet = ss.getSheetByName("문자발송");
    var personalSheet = ss.getSheetByName("개인_알림");
    
    var historyLogs = [];
    
    // 1. 문자발송 시트 스캔 (최근 500줄 스캔하여 특정 회원 찾기)
    if (smsSheet) {
      var smsLast = smsSheet.getLastRow();
      if (smsLast >= 2) {
        var startRow = Math.max(2, smsLast - 500);
        var numRows = smsLast - startRow + 1;
        var smsData = smsSheet.getRange(startRow, 1, numRows, 6).getDisplayValues();
        
        for (var i = smsData.length - 1; i >= 0; i--) {
          var row = smsData[i];
          var rowPhone = String(row[2] || "").replace(/[^0-9]/g, "");
          if (rowPhone === cleanPhone || (cleanPhone.length >= 8 && rowPhone.endsWith(cleanPhone.substring(cleanPhone.length - 8)))) {
            var rawContent = String(row[4] || "").trim();
            // 장해 진단서 문구가 들어있거나, 지나치게 짧고 온전한 구두점 없이 끝나는 경우 이력 매칭에서 과감히 제외
            if (rawContent.indexOf("장해 진단") !== -1 || (rawContent.length < 50 && !rawContent.endsWith(".") && !rawContent.endsWith("!") && !rawContent.endsWith("❤️") && !rawContent.endsWith("~"))) {
              continue;
            }
            historyLogs.push({
              date: row[0],
              channel: "일반 SMS 문자",
              category: row[3],
              content: rawContent,
              status: row[5]
            });
          }
          if (historyLogs.length >= 4) break;
        }
      }
    }
    
    // 2. 개인쪽지(인앱 알림) 시트 스캔 (최근 500줄 스캔하여 특정 회원 찾기)
    if (personalSheet && historyLogs.length < 4) {
      var personalLast = personalSheet.getLastRow();
      if (personalLast >= 2) {
        var startRow = Math.max(2, personalLast - 500);
        var numRows = personalLast - startRow + 1;
        var personalData = personalSheet.getRange(startRow, 1, numRows, 8).getDisplayValues();
        
        for (var i = personalData.length - 1; i >= 0; i--) {
          var row = personalData[i];
          var rowPhone = String(row[1] || "").replace(/[^0-9]/g, "");
          if (rowPhone === cleanPhone || (cleanPhone.length >= 8 && rowPhone.endsWith(cleanPhone.substring(cleanPhone.length - 8)))) {
            var rawContent = String(row[5] || "").trim();
            if (rawContent.indexOf("장해 진단") !== -1 || (rawContent.length < 50 && !rawContent.endsWith(".") && !rawContent.endsWith("!") && !rawContent.endsWith("❤️") && !rawContent.endsWith("~"))) {
              continue;
            }
            historyLogs.push({
              date: row[6],
              channel: "인앱 쪽지(알림)",
              category: row[3],
              content: rawContent,
              status: "완료"
            });
          }
          if (historyLogs.length >= 6) break; // 최대 6건 취합
        }
      }
    }
    
    if (historyLogs.length === 0) {
      return "최근 발송 이력 없음 (신규 케어 대상 회원입니다.)";
    }
    
    // 시간순 정렬 (오래된 것 -> 최신 것)
    historyLogs.sort(function(a, b) {
      return new Date(a.date) - new Date(b.date);
    });
    
    // 이력 텍스트화
    var textArr = [];
    historyLogs.forEach(function(log, idx) {
      textArr.push("[" + (idx + 1) + "] " + log.date + " | 발송채널: " + log.channel + " | 분류: " + log.category + " | 발송상태: " + log.status + "\n└ 메시지내용: " + log.content);
    });
    
    return textArr.join("\n\n");
  } catch (e) {
    Logger.log("🚨 getMemberMessageHistory 오류: " + e.toString());
    return "발송 이력 조회 중 오류 발생 (무결한 신규 멘트 작성 필요)";
  }
}

/**
 * 🤖 [v64.40] 지니 웰니스 AI 문자/쪽지 통합 생성기 (Gemini API & 히스토리 지능 융합형)
 * 회원의 상황(이름, 남은횟수, 경과일수, 유형, 번호, 발송채널)을 받아, 과거 발송 이력을 읽은 제미나이가
 * 회원이 스팸처럼 느끼지 않고 따뜻한 안부를 체감하도록 최적화된 맞춤형 복귀 권유/안부 문구를 실시간 창조합니다.
 */
function generateWellnessAiSms(name, remain, elapsedDays, type, phone, targetChannel, memberClassInfo, hasNoAttendance, expireDateParam) {
  var cleanName = name.replace(/\d{4}$/, ""); // 이름 끝 숫자 제거
  var recentHistoryText = getMemberMessageHistory(phone); // 최근 3~5건의 문자+쪽지 발송 내역 취합
  
  var channelInfo = targetChannel === "noti" ? "인앱 쪽지(알림)" : "오프라인 SMS 일반 문자";
  
  var originalClass = memberClassInfo || "general";
  var classInfo = memberClassInfo || "general";
  if (hasNoAttendance) {
    classInfo = "new_no_record";
  }
  
  // [v65.00] 실시간 이용 만료일(유효기간) 동적 조회 (넘겨받은 파라미터가 있다면 우선 적용하여 버그 예방!)
  var expireDateStr = "기록없음 (미입력)";
  if (expireDateParam && String(expireDateParam).trim() !== "" && String(expireDateParam) !== "미정" && String(expireDateParam) !== "기록없음 (미입력)") {
    expireDateStr = String(expireDateParam).trim();
  } else {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    if (regSheet && phone) {
      var regData = regSheet.getDataRange().getDisplayValues();
      var regCols = getRegColumnIndices(regSheet);
      var cleanPhone = phone.replace(/[^0-9]/g, "");
      var lastExpire = new Date(0);
      for (var i = 1; i < regData.length; i++) {
        var rowPhone = String(regData[i][regCols.phone] || "").replace(/[^0-9]/g, "");
        if (rowPhone === cleanPhone || (cleanPhone.length >= 8 && rowPhone.endsWith(cleanPhone.substring(cleanPhone.length - 8)))) {
          var rowStatus = String(regData[i][regCols.status] || "").trim();
          if (rowStatus === "진행중" || rowStatus === "진행 중") {
            var rowExpireStr = regData[i][regCols.expire];
            if (rowExpireStr) {
              var expDate = new Date(rowExpireStr);
              if (expDate > lastExpire) lastExpire = expDate;
            }
          }
        }
      }
      if (lastExpire.getTime() > 0) {
        expireDateStr = Utilities.formatDate(lastExpire, "GMT+9", "yyyy-MM-dd");
      }
    }
  }
  
  // 🚨 [회원권 성향별 맞춤 가입 주입 - v64.70 신규 미출석 무기록 가이드 탑재]
  var classDesc = "";
  if (classInfo === "therapy") {
    classDesc = "★[회원 성향 맞춤] 이 회원은 점핑 운동을 하지 않고 '원적외선 테라피(반신욕/테라피룸 편백 힐링)'만을 이용하던 힐링 전용 회원입니다. 절대 격렬한 점핑 운동 복귀나 트램폴린 뛰기 등을 언급하지 마시고, 편안하고 따뜻하게 땀 빼며 쉴 수 있는 편백 원적외선 테라피실의 안부와 여유를 핵심 주제로 삼아 작성하십시오.";
  } else if (classInfo === "jumping") {
    classDesc = "★[회원 성향 맞춤] 이 회원은 테라피를 이용하지 않고 오직 '트램폴린 점핑 운동'만을 열심히 하던 운동 전용 회원입니다. 원적외선 테라피나 반신욕실 언급은 완전히 배제하시고, 신나는 음악과 강력한 칼로리 소모, 체력 회복의 '점핑 운동 에너지'를 중심으로 복귀를 독려하십시오.";
  } else if (classInfo === "complex") {
    classDesc = "★[회원 성향 맞춤] 이 회원은 신나는 '트램폴린 점핑 운동'과 편안한 '원적외선 테라피 힐링'을 모두 즐기던 복합(하이브리드) 회원입니다. 점핑 운동의 활기찬 땀방울과 테라피의 따뜻한 피로 회복, 두 가지 매력을 조화롭고 자연스럽게 녹여서 감성 메시지를 조립하십시오.";
  } else if (classInfo === "new_no_record") {
    var originalDesc = "신나고 활기찬 점핑 운동과 따뜻한 테라피 힐링";
    if (originalClass === "therapy") originalDesc = "따뜻한 편백 원적외선 테라피실의 아늑한 힐링과 여유";
    else if (originalClass === "jumping") originalDesc = "신나는 음악과 활력 넘치는 트램폴린 점핑 운동";
    
    classDesc = "★[회원 성향 맞춤] 이 회원은 노형점핑클럽에 등록은 되어 있으나, 시스템 상에 과거 출석 기록이 아예 존재하지 않는 신규 또는 이관(마이그레이션) 무기록 회원입니다. '이전에 오셨을 때 편안하셨길 바란다'거나 '다시 온다'는 식의 유경험 뉘앙스 표현은 절대 쓰지 마십시오. 대신 '등록하신 이후 아직 클럽에서 첫 걸음을 내딛지 못하셨거나 첫 운동을 어색함으로 주저하고 계신 것 같다'며, 특별히 이 회원이 보유한 이용권 성향인 [" + originalDesc + "]을 따뜻하게 소개하고, '코치와 함께하는 첫 번째 웰니스 여정'을 설렘 가득하고 따뜻하게 열어드리는 웰커밍 메시지(환대 및 첫 방문 유도)로 독려하십시오.";
  } else {
    classDesc = "★[회원 성향 맞춤] 이 회원은 일반 점핑클럽 회원입니다. 활력 넘치는 점핑 운동 또는 따뜻한 테라피 힐링 안부를 편안하고 균형감 있게 터치하며 소통하십시오.";
  }
  
  var systemInstruction = "당신은 제주 노형점핑클럽의 친근하고 열정적인 최고의 '웰니스 코치'입니다. " +
                          "마을 주민(회원)들의 건강을 진심으로 걱정하고 다시 클럽으로 복귀하도록 이끄는 따뜻하고 감성 넘치는 메시지를 작성해 주세요. " +
                          "글자수는 " + channelInfo + " 발송용이므로 150자~250자 내외로 간결하게 작성하고, 격려와 친근함을 담아 이모티콘(😊, ❤️, 🏃‍♀️, 🔥)을 적절히 융합해 주세요. 절대 딱딱하거나 기계적인 말투를 쓰지 마세요.\n\n" +
                          "★가장 중요한 비즈니스 및 작문 규칙 (필수 준수)★\n" +
                          "1. 📅 이용 기한(만료일)과 📊 잔여 횟수를 한눈에 잘 들어오게 표기하여 강조해야 합니다.\n" +
                          "   - 예: '📅 이용 기한: ~ O월 O일까지' 및 '📊 잔여 횟수: O회'\n" +
                          "   - 잔여 횟수만 언급하면 기한을 잊어버리고 안 오는 경우가 많으므로, 반드시 '**이용 기한인 [만료일]까지 남은 [횟수]회를 소진하셔야 하며, 기한이 지나면 소멸된다**'는 점을 아주 친근하고 따뜻하면서도 확실하게 각인시켜 주십시오. 기한 내에 다 쓰실 수 있도록 서둘러 늦지 않게 방문을 독려하는 멘트를 필수 주입하십시오.\n\n" +
                          "2. 분류 유형이 '장기미방문' (결석 중)인 경우:\n" +
                          "   - 이 회원은 아직 수강권이 마감되지 않았고 남은 회원권 횟수가 유효하게 보존되어 있는 상태입니다.\n" +
                          "   - 따라서 '원적외선 테라피 무료 체험', '깜짝 소생 보너스 복구', '추가 보너스 이용권 제공' 등 어떠한 무료 서비스나 추가 보상 혜택도 절대 제시하거나 언급하지 마십시오.\n" +
                          "   - 오직 '회원님이 보유하고 계신 소중한 남은 회원권(잔여 횟수)이 아깝게 기한 만료로 소멸되기 전에 어서 클럽에 복귀하여 유효기간 안에 다 쓰시라'고 복귀를 따뜻하게 권유하고 소모를 독려하는 내용만으로 메시지를 작성하십시오.\n\n" +
                          "3. 분류 유형이 '복귀권유' (수강 마감 상태)인 경우 (마감된 회원):\n" +
                          "   - [중요: 기간별 문조 차별화]\n" +
                          "     * 마감 경과일이 최근(1일~14일)인 경우: '최근에 이용권이 마감되었다'며 빠른 복귀와 안부를 유도하십시오.\n" +
                          "     * 마감 경과일이 중기(15일~60일)인 경우: '이용권이 마감되고 한동안 뵙지 못했는데 그간 잘 지내셨는지 궁금하다'는 안부로 대화를 유도하십시오.\n" +
                          "     * 마감 경과일이 장기(61일 이상 또는 무기록 999일)인 경우: '마지막으로 뵌 지도 벌써 오랜 시간이 흘렀다'거나 '오랫동안 소식을 여쭙지 못해 보고 싶었다'는 깊은 감성 안부를 담아 문장을 시작하고, 절대 최근에 마감되었다고 말하지 마십시오.\n" +
                          "   - [중요: 잔여 횟수에 따른 혜택 통제 규칙]\n" +
                          "     * 잔여 횟수가 7회 미만(0회~6회)인 경우: 잔여 횟수 소생이나 보너스 복구 등의 언급을 절대 하지 마십시오. 5회 이하 등의 소량은 소생 멘트 없이 '체력 루틴 복귀 및 테라피 힐링 안부' 중심으로 재등록만 따뜻하게 유치하십시오.\n" +
                          "     * 잔여 횟수가 7회 이상인 경우: 남은 횟수가 많아 안타깝다는 아쉬움은 따뜻하게 언급(예: '이용권 마감 당시 쓰지 못하셨던 잔여 횟수들이 꽤 많이 남았어서 마음이 아프다')하되, '이걸 다 살려주겠다'는 100% 직접적 복구 표현은 나중에 분란의 소지가 있으므로 절대 금지하십시오! 대신 '이번에 다시 오셔서 시작하시면 코치가 안타까운 마음을 채워드릴 수 있게 넉넉한 보너스 횟수를 챙겨서 보태드리겠다'는 식의 넉넉하고 두루뭉술한 혜택 보너스 표현으로 작성하십시오.\n\n" +
                          "4. 제공되는 [최근 메시지 발송 이력]을 반드시 분석하여, 이전과 겹치지 않는 신선한 멘트를 창출하십시오.\n" +
                          "5. " + classDesc;
  
  // 🛡️ [원장님 스마트 가드] 출석 기록이 아예 없어 999일로 기입되는 회원에 대한 AI 프롬프트 정밀 가드
  var elapsedDaysPromptVal = (elapsedDays === 999) ? "기록 없음 (오랫동안 미방문 상태)" : elapsedDays + "일";
  if (classInfo === "new_no_record") {
    elapsedDaysPromptVal = "기록 없음 (신규/무기록 회원)";
  }
  
  var prompt = "■ 회원 상황 정보\n" +
               "- 회원 이름: " + cleanName + "\n" +
               "- 회원권 잔여 횟수: " + remain + "회\n" +
               "- 회원권 이용 기한(만료일): " + expireDateStr + "\n" +
               "- 미출석/만료 경과일: " + elapsedDaysPromptVal + "\n" +
               "- 분류 유형: " + type + " (" + (type === "장기미방문" ? "결석 중 (남은 회원권 소진 독려 대상)" : "수강 마감 상태 (재등록 유치 대상)") + ")\n" +
               "- 회원 이용 성향: " + classInfo + " (" + (classInfo === "therapy" ? "테라피 전용" : classInfo === "jumping" ? "점핑 운동 전용" : classInfo === "complex" ? "점핑 및 테라피 복합" : classInfo === "new_no_record" ? "신규 가입 무기록 회원" : "일반 회원") + ")\n" +
               "- 발송 예정 채널: " + channelInfo + "\n\n" +
               "■ [중요] 최근 메시지 발송 이력 (시공간 순)\n" +
               recentHistoryText + "\n\n" +
               "위 상황 정보와 최근 발송 히스토리를 꿰뚫어 보고, 이 회원에게 다음 단계로 전달할 가장 자연스럽고 마음을 끄는 최고의 넥스트 " + channelInfo + " 문장을 새로 작성해 주세요.\n" +
               "※ 필수 준수:\n" +
               "- 기한인 " + expireDateStr + "과 남은 횟수 " + remain + "회를 글에 알맞게 녹여내고, 이 기한 내에 다 쓰지 않으면 횟수가 소멸된다는 중요 내용을 확실히 일러두십시오.\n" +
               "- 분류 유형이 '장기미방문'인 경우 무료 혜택이나 보너스 선물 언급을 100% 배제하고 '남은 회원권을 기간 내에 쓰시라'는 취지만 따뜻하게 작성하십시오.\n" +
               "- 분류 유형이 '복귀권유'이고 경과일이 오래된 장기 미등록 회원의 경우 절대 '며칠 전 이용권이 마감되었다'고 쓰지 말고 오랜 소식을 여쭙는 감성 멘트를 쓰십시오.\n" +
               "- 잔여 횟수가 7회 미만이면 횟수 살려주겠다는 보너스 복구 제안을 절대 하지 마십시오.\n" +
               "- 잔여 횟수가 7회 이상이면 '남은 횟수가 많아 마음 아팠다'~과 같이 언급하고 '다시 시작하시면 넉넉히 보너스 횟수를 챙겨 보태겠다'며 절대 100% 다 살려준다는 명시적 단어 없이 은근하고 포근하게 작성하십시오.\n" +
               "- 경과일이 '기록 없음'이거나 무기록인 경우 구체적인 일수(예: 999일 등)를 절대 문장에 언급하지 마십시오.\n" +
               "- 회원의 이용 성향(" + classInfo + ")에 집중하여 테라피 전용 회원은 점핑 운동 언급을 완전히 배제하고, 점핑 전용 회원은 원적외선 테라피 언급을 완전히 배제하고, new_no_record는 첫 방문 환대 언급으로 작성하십시오.";
  
  var aiMsg = callGeminiBackend(prompt, systemInstruction);
  
  // 만약 API 호출 실패 시 감성 템플릿(Fallback 백업)으로 안전 복구
  if (!aiMsg) {
    if (classInfo === "new_no_record") {
      var newThemeText = "신나고 짜릿한 점핑 운동과 몸을 사르르 녹여주는 따뜻한 테라피 힐링";
      if (originalClass === "therapy") {
        aiMsg = cleanName + "회원님, 노형점핑클럽입니다. 😊 저희 클럽에 소중한 첫 등록을 해주신 이후로, 아직 편안한 힐링실에서 뵙지 못한 것 같아 설레는 마음 반, 걱정스러운 마음 반으로 안부 전합니다. 💖 혹시 첫 방문이 낯설거나 예약이 어려우셨다면 부담은 완전히 내려놓고 편하게 들러주세요! 지치고 찌푸둥한 몸을 따뜻하게 충전하실 수 있게 코치가 정성껏 에스코트해 드릴게요. 웰니스 코치와 함께 첫 힐링 타임을 시작해 보시는 건 어떨까요? 기다리고 있을게요! 🌿❤️";
      } else if (originalClass === "jumping") {
        aiMsg = cleanName + "회원님, 노형점핑클럽입니다. 😊 신나고 신선한 활력을 약속하는 저희 클럽에 등록해주신 이후, 아직 첫 수업을 시작하지 못하신 것 같아 설레는 마음 반, 주저하시는 건 아닌가 걱정되는 마음 반으로 응원 전합니다. 💖 처음이라 어색하거나 망설여지신다면 걱정 마세요! 기초부터 아주 쉽고 재미있게 뛰실 수 있도록 코치가 확실히 케어해 드릴게요. 이번 주에 신나는 음악 속에 첫 걸음을 떼어보시는 건 어떨까요? 파이팅! 🔥💪";
      } else {
        aiMsg = cleanName + "회원님, 노형점핑클럽입니다. 😊 저희 클럽에 소중한 첫 발걸음을 딛으신 이후로, 아직 한 번도 뵙지 못한 것 같아 설레는 마음 반, 걱정스러운 마음 반으로 조심스레 안부 전합니다. 💖 혹시 첫 운동이 낯설거나 시작이 망설여지셨다면 부담은 완전히 내려놓고 놀러오세요! " + newThemeText + "을 코치가 아주 친절하고 재미있게 경험하실 수 있도록 에스코트해 드릴게요. 함께 건강한 첫 걸음을 웰니스 코치와 시작해 보시는 건 어떨까요? 기다리고 있겠습니다! ❤️";
      }
    } else if (type === "장기미방문") {
      var elapsedTextFallback = (elapsedDays === 999) ? "오랫동안 소식이 닿지 않아 무척 궁금했답니다." : "뵙지 못한 지 벌써 " + elapsedDays + "일이나 지났네요! 😢";
      
      var defaultTheme = "가벼운 점핑 운동이나 따뜻한 테라피";
      if (classInfo === "therapy") defaultTheme = "편안하고 따뜻한 원적외선 테라피실 힐링";
      else if (classInfo === "jumping") defaultTheme = "신나고 에너지 넘치는 점핑 운동";
      
      aiMsg = cleanName + "회원님, 노형점핑클럽입니다. 😊 클럽에서 " + elapsedTextFallback + " 많이 바쁘시더라도 " + defaultTheme + "로 다시 건강 리듬을 회복해보시는 건 어떨까요? 소중한 회원님의 잔여 회원권이 아직 " + remain + "회나 남았습니다. 아깝게 소멸되기 전에 얼른 오셔서 함께 건강을 가꾸어보아요! ❤️";
    } else {
      // 1. 복귀 권유(마감된 회원) 기간별 문조 설정
      var periodText = "최근에 이용권(수강권)이 마감되었네요. 혹시 깜빡 잊고 계셨던 건 아닌지 궁금해요! 😊";
      if (elapsedDays >= 15 && elapsedDays <= 60) {
        periodText = "이용권이 마감되고 한동안 뵙지 못했는데, 그동안 건강하게 잘 지내셨는지 문득 걱정되고 궁금해졌답니다. 😊";
      } else if (elapsedDays >= 61 || elapsedDays === 999) {
        periodText = "마지막으로 뵌 지도 벌써 오랜 시간이 흘렀네요! 오랫동안 소식을 여쭙지 못해 늘 마음 한켠으로 보고 싶었답니다. 😊";
      }
      
      // 2. 이용 성향별 문구 테마 설정 (테라피 전용, 점핑 전용, 복합형)
      var themeText = "가벼운 운동과 따뜻한 테라피실에서 건강 에너지를 다시 충전해보시는 건 어떨까요?";
      if (classInfo === "therapy") {
        themeText = "따뜻한 편백나무 원적외선 테라피실에서 개운하게 땀 빼며 힐링과 여유를 다시 만끽해보시는 건 어떨까요?";
      } else if (classInfo === "jumping") {
        themeText = "신나는 음악과 함께 트램폴린을 뛰며 스트레스도 날리고 활기찬 운동 리듬을 다시 회복해보시는 건 어떨까요?";
      } else if (classInfo === "complex") {
        themeText = "신나게 점핑 운동을 뛰며 땀을 쫙 흘리고, 따뜻한 원적외선 테라피실에서 온몸을 사르르 녹여보시는 건 어떨까요?";
      }

      // 3. 7회 기준 혜택 분기 조립
      if (remain >= 7) {
        aiMsg = cleanName + "회원님, 노형점핑클럽입니다. " + periodText + " 마감될 당시에 사용하지 못하셨던 아까운 잔여 횟수가 꽤 많이 남았던 게 계속 마음에 걸렸어요. 😢 이번에 다시 오셔서 기분 좋게 새로 시작하시고 " + themeText.replace(" 어떨까요?", "") + "면, 아쉬우셨던 마음까지 채워드릴 수 있도록 제가 넉넉한 보너스 횟수를 기분 좋게 선물로 보태드리겠습니다. 어서 편하게 들러주세요! ❤️";
      } else {
        if (remain === 0) {
          aiMsg = cleanName + "회원님, 노형점핑클럽입니다. " + periodText + " 겨우 다져놓았던 건강한 리듬과 " + (classInfo === "therapy" ? "테라피 힐링 루틴" : "점핑 운동 루틴") + "이 이대로 멈춰서기엔 너무 아깝잖아요. 😢 " + themeText + " 가벼운 마음으로 다시 클럽에 들러주시면 코치가 최선을 다해 정성껏 웰니스 케어를 도와드리겠습니다. 얼른 다시 뵙고 싶어요! ❤️";
        } else {
          aiMsg = cleanName + "회원님, 노형점핑클럽입니다. " + periodText + " 아깝게 이용권 기한이 다해 멈췄던 루틴이 너무 아쉽네요. 😢 이번 기회에 다시 클럽으로 복귀하셔서 " + themeText + " 코치가 최선을 다해 건강 파트너가 되어드릴 테니, 부담 없이 얼굴 보여주러 오세요! ❤️";
        }
      }
    }
  }
  
  return aiMsg;
}

/**
 * 📱 [v64.40] 특정 회원의 최근 모바일 앱 활동(로그인/다이어리 등) 여부 검사기
 * '일일_활동_기록' 시트를 스캔하여 최근 limitDays 일 동안 해당 회원의 활동 내역이 1건이라도 있는지 판단합니다.
 */
function hasMemberActiveInApp(phone, limitDays) {
  try {
    var cleanPhone = String(phone || "").replace(/[^0-9]/g, "");
    if (!cleanPhone) return false;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("일일_활동_기록");
    if (!sheet) return false;
    
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    
    // 최근 800줄만 스캔하여 속도 최적화
    var startRow = Math.max(2, lastRow - 800);
    var numRows = lastRow - startRow + 1;
    var data = sheet.getRange(startRow, 1, numRows, 3).getDisplayValues();
    
    var now = new Date();
    var limitMs = limitDays * 24 * 60 * 60 * 1000;
    
    for (var i = data.length - 1; i >= 0; i--) {
      var row = data[i];
      var rowPhone = String(row[1] || "").replace(/[^0-9]/g, "");
      
      if (rowPhone === cleanPhone || (cleanPhone.length >= 8 && rowPhone.endsWith(cleanPhone.substring(cleanPhone.length - 8)))) {
        var logDateStr = row[0];
        if (logDateStr) {
          var logDate = new Date(logDateStr);
          var diffMs = now.getTime() - logDate.getTime();
          if (diffMs <= limitMs) {
            return true; // 최근 limitDays 일 내에 앱 활동이 존재함!
          }
        }
      }
    }
    return false;
  } catch (e) {
    Logger.log("🚨 hasMemberActiveInApp 오류: " + e.toString());
    return false;
  }
}

/**
 * 🤖 [v64.40] 수동 발송용 AI 메시지 초안 생성기 (Gemini API & 히스토리 융합형)
 * 원장님이 관리자 에디터에서 수동으로 개인쪽지/공지 등을 작성할 때 AI가 상황 정보 및 히스토리를
 * 100% 수집하여 감성 넘치는 초정밀 맞춤 초안을 기획해 줍니다.
 */
function generateAiDraftForManualMessage(payload) {
  try {
    var rawPhone = String(payload.phone || "");
    var name = String(payload.name || "");
    var memo = String(payload.memo || "").trim();
    var targetType = String(payload.targetType || "personal"); // personal 또는 global
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (targetType === "global") {
      // 1. 단체 공지 사항 작문 시나리오
      var systemInstruction = "당신은 제주 노형점핑클럽의 열정적인 '웰니스 코치'이자 친근한 건강 메신저입니다. " +
                              "클럽 전체 모험가(회원)들에게 공지할 감성 넘치고 모험심을 자극하는 멋진 공지문 본문을 써주세요. " +
                              "어투는 다정하고 활기차게 이모티콘을 활용해 주세요.";
      var prompt = "웰니스 코치의 공지 요지 지시사항: [" + (memo || "일반 유쾌한 건강 다이어리 피드백 및 안부 인사") + "]\n\n" +
                   "위 지시사항을 멋지게 가공하여 전체 주민들이 기쁜 마음으로 동참할 만한 150자 내외의 세련된 공지문 본문을 한 판 써주세요.";
      var draft = callGeminiBackend(prompt, systemInstruction);
      return { success: true, draft: draft || "공지 사항을 적어보세요! ❤️" };
    }
    
    // 2. 개별 1:1 쪽지 작문 시나리오 (프로필 + 최근 발송 히스토리 자동 추출)
    var cleanPhone = rawPhone.replace(/[^0-9]/g, "");
    var regSheet = ss.getSheetByName("등록 현황");
    
    var remain = 0;
    var elapsedDays = 0;
    var status = "만료";
    
    if (regSheet && cleanPhone) {
      var regData = regSheet.getDataRange().getDisplayValues();
      var regCols = getRegColumnIndices(regSheet);
      
      // 해당 폰 번호의 만료일과 횟수 스캔
      var now = new Date();
      var lastExpire = new Date(0);
      var isActive = false;
      
      for (var i = 1; i < regData.length; i++) {
        var rowPhone = String(regData[i][regCols.phone] || "").replace(/[^0-9]/g, "");
        if (rowPhone === cleanPhone || (cleanPhone.length >= 8 && rowPhone.endsWith(cleanPhone.substring(cleanPhone.length - 8)))) {
          var rowStatus = String(regData[i][regCols.status] || "").trim();
          var rowRemain = parseInt(regData[i][regCols.remain]) || 0;
          var rowExpireStr = regData[i][regCols.expire];
          
          remain += rowRemain;
          if (rowStatus === "진행중" || rowStatus === "진행 중") isActive = true;
          if (rowExpireStr) {
            var expDate = new Date(rowExpireStr);
            if (expDate > lastExpire) lastExpire = expDate;
          }
        }
      }
      
      status = isActive ? "장기미방문" : "복귀권유";
      if (lastExpire.getTime() > 0 && lastExpire < now) {
        var diffMs = now.getTime() - lastExpire.getTime();
        elapsedDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
      } else {
        // 출석 기록 스캔 (결석 일수 구하기)
        var logSheet = ss.getSheetByName("출석기록");
        if (logSheet) {
          var logData = logSheet.getDataRange().getDisplayValues();
          var logCols = getAttendanceColumnIndices(logSheet);
          var lastAttendance = new Date(0);
          for (var j = 1; j < logData.length; j++) {
            var logPhone = String(logData[j][logCols.phone] || "").replace(/[^0-9]/g, "");
            if (logPhone === cleanPhone || (cleanPhone.length >= 8 && logPhone.endsWith(cleanPhone.substring(cleanPhone.length - 8)))) {
              var logDate = new Date(logData[j][logCols.date]);
              if (logDate > lastAttendance) lastAttendance = logDate;
            }
          }
          if (lastAttendance.getTime() > 0) {
            var diffMs = now.getTime() - lastAttendance.getTime();
            elapsedDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
          }
        }
      }
    }
    
    var recentHistoryText = getMemberMessageHistory(cleanPhone);
    var cleanName = name.replace(/\d{4}$/, ""); // 이름 끝 숫자 제거
    
    // [v65.00] 만료일 포맷 변환
    var formattedExpireDate = "기록없음 (미입력)";
    if (lastExpire && lastExpire.getTime() > 0) {
      formattedExpireDate = Utilities.formatDate(lastExpire, "GMT+9", "yyyy-MM-dd");
    }
    
    var systemInstruction = "당신은 제주 노형점핑클럽의 다정한 '웰니스 코치'이자 최고의 헬스 코치입니다. " +
                            "웰니스 코치가 특정 회원에게 1:1로 직접 보낼 '감성 케어 편지/쪽지'의 명품 본문을 작성해야 합니다.\n" +
                            "제공되는 [최근 메시지 발송 이력]과 [웰니스 코치의 작문 메모]를 100% 정밀 분석하여, " +
                            "중복 멘트를 완벽 방어하고 회원과 1:1로 친근하게 속삭이는 감동적인 다이어리 피드백/안부 쪽지를 지어내십시오.\n\n" +
                            "⚠️ 가장 중요한 비즈니스 및 작문 규칙 (필수 준수):\n" +
                            "1. 📅 이용 기한(만료일)과 📊 잔여 횟수를 문장 내에 잘 들어오게 표기하여 강조해야 합니다.\n" +
                            "   - 예: '📅 이용 기한: ~ O월 O일까지' 및 '📊 잔여 횟수: O회'\n" +
                            "   - 잔여 횟수만 언급하면 기한을 잊어버리고 안 오는 경우가 많으므로, 반드시 '**이용 기한인 [만료일]까지 남은 [횟수]회를 소진하셔야 하며, 기한이 지나면 소멸된다**'는 점을 아주 친근하고 따뜻하면서도 확실하게 각인시켜 주십시오. 기한 내에 다 쓰실 수 있도록 서둘러 늦지 않게 방문을 독려하는 멘트를 필수 주입하십시오.\n" +
                            "2. 최근 발송 히스토리에 있는 문자 내용이나 끊긴 예전 문구를 절대로 그대로 복사하거나 뒤이어 쓰지 마십시오. 항상 기승전결이 완벽히 마감되는 새로운 편지를 창작해야 합니다.";
    
    var prompt = "■ 수신인 회원 정보\n" +
                 "- 회원 이름: " + cleanName + "\n" +
                 "- 잔여 수강 횟수: " + remain + "회\n" +
                 "- 회원권 이용 기한(만료일): " + formattedExpireDate + "\n" +
                 "- 미출석/만료 경과일: " + elapsedDays + "일 (" + status + " 상태)\n" +
                 "- 웰니스 코치의 직접 지시사항(작문 요지): [" + (memo || "다정하고 편안하게 안부와 건강 챙기기") + "]\n\n" +
                 "■ 최근 발송 히스토리\n" +
                 recentHistoryText + "\n\n" +
                 "⚠️ [작문 지침]\n" +
                 "1. 위 정보와 웰니스 코치의 지시 메모를 기반으로 회원의 가슴을 뭉클하게 할 150자~200자 내외의 명품 1:1 쪽지 초안 본문을 작성해 주세요.\n" +
                 "2. 기한인 " + formattedExpireDate + "과 남은 횟수 " + remain + "회를 텍스트 중간에 인용하고, 기한 내에 모두 쓰셔야 소멸되지 않는다는 점을 명확히 리마인드하십시오.\n" +
                 "3. 절대로 히스토리에 들어있는 기존 메시지 문구(특히 끊긴 문구)를 베끼거나 복제하지 마십시오.\n" +
                 "4. 메시지 본문은 절대로 중간에 뚝 끊기면 안 되며, 반드시 온전한 마침표('.') 또는 이모티콘과 함께 완결된 종결어미('~요.', '~입니다.', '❤️')로 확실하게 끝마쳐야 합니다.";
    
    var apiRes = callGeminiBackendWithDetails(prompt, systemInstruction);
    var draft = "";
    if (apiRes.success) {
      draft = apiRes.text;
    } else {
      // 🚨 [진단용 핫패치 v64.45] API 호출의 날것 그대로인 에러 전문을 출력!
      var apiKeyRes = getGeminiApiKey();
      var rawKey = apiKeyRes.key || "";
      var keySnippet = (rawKey.length > 5) ? (rawKey.substring(0, 5) + "..." + rawKey.substring(rawKey.length - 3)) : "비어있음/추출안됨";
      
      draft = "[⚠️ 지니 비서 실시간 장해 진단 보고서]\n" +
              "1. 환경설정 시트 API Key 로드 결과: " + (apiKeyRes.success ? "성공 ✅" : "실패 ❌") + "\n" +
              "2. 감지된 API Key 상태: " + keySnippet + " (길이: " + rawKey.length + "글자)\n" +
              "3. 백엔드가 감지한 진짜 오류 원인:\n➡️ " + apiRes.error + "\n\n" +
              "👉 임시 쪽지: " + cleanName + "님, 이용 기한(" + formattedExpireDate + ") 내에 남은 회원권(" + remain + "회)을 소중하게 다 쓰실 수 있도록 서둘러 자주 방문해 주세요! 늦기 전에 클럽에서 뵙겠습니다. ❤️";
    }
    return { success: true, draft: draft };
    
  } catch (e) {
    Logger.log("🚨 generateAiDraftForManualMessage 오류: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

/**
 * ==========================================
 * 🌦️ 실시간 날짜/날씨 컨트롤 및 BGM 설정 API (Cyworld Nostalgia)
 * ==========================================
 */
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

// [이동 완료] resolveSunoUrl 및 resolveSingleSunoUrl 함수는 Bridge_Login.gs로 이동되었습니다.

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
 * 🌵 오아시스 돌발 게시판 (소통/다짐 스레드) API
 * ==========================================
 */
function getOasisPosts() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("오아시스_글") || ss.insertSheet("오아시스_글");
    if (sheet.getLastRow() < 1) {
      sheet.appendRow(["날짜", "시간", "연락처", "작성자", "제목", "내용", "구분", "칭찬대상", "공감수", "공감클릭연락처들"]);
      sheet.setFrozenRows(1);
      return [];
    }
    
    // [v55.0 컬럼 보정] 기존 8개 컬럼 구조일 경우 10개 컬럼으로 자동 헤더 확장
    if (sheet.getLastColumn() < 10) {
      var headerRange = sheet.getRange(1, 1, 1, 10);
      headerRange.setValues([["날짜", "시간", "연락처", "작성자", "제목", "내용", "구분", "칭찬대상", "공감수", "공감클릭연락처들"]]);
    }

    var data = sheet.getDataRange().getDisplayValues();
    var posts = [];
    for (var i = 1; i < data.length; i++) {
      posts.push({
        id: i + 1, // 시트 행 인덱스와 동기화 (1-based row index)
        date: data[i][0],
        time: data[i][1],
        phone: data[i][2],
        author: data[i][3],
        title: data[i][4],
        content: data[i][5],
        category: data[i][6] || "일반",
        targetMember: data[i][7] || "",
        hearts: Number(data[i][8] || 0),
        heartPhones: data[i][9] || ""
      });
    }
    return posts.reverse(); // 최신글이 상단으로
  } catch (e) {
    return [];
  }
}

function submitOasisPost(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("오아시스_글") || ss.insertSheet("오아시스_글");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["날짜", "시간", "연락처", "작성자", "제목", "내용", "구분", "칭찬대상", "공감수", "공감클릭연락처들"]);
      sheet.setFrozenRows(1);
    }
    
    // [v55.0 컬럼 보정]
    if (sheet.getLastColumn() < 10) {
      var headerRange = sheet.getRange(1, 1, 1, 10);
      headerRange.setValues([["날짜", "시간", "연락처", "작성자", "제목", "내용", "구분", "칭찬대상", "공감수", "공감클릭연락처들"]]);
    }

    var now = new Date();
    var dateStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var timeStr = Utilities.formatDate(now, "GMT+9", "HH:mm:ss");
    var phone = String(payload.phone || "").replace(/[^0-9]/g, "");
    var category = payload.category || "일반";
    var targetMember = payload.targetMember || "";
    
    sheet.appendRow([
      dateStr,
      timeStr,
      "'" + phone,
      payload.author,
      payload.title,
      payload.content,
      category,
      targetMember,
      0,  // 초기 공감수
      ""  // 초기 공감클릭연락처들
    ]);
    
    var awardMessage = "";
    var expAwarded = 0;
    var statType = "def"; // 기본적으로 칭찬/수호 다짐은 회복력(def)으로 연동!
    var itemTitle = "";
    
    var qTitle = payload.questTitle || "";
    var qScore = 15;
    if (category === "돌발") {
      try {
        var questStatus = getActiveQuestStatus(phone, ss, null, payload.author);
        if (questStatus && questStatus.todayQuest) {
          if (!qTitle) qTitle = questStatus.todayQuest.title;
          if (questStatus.todayQuest.score) {
            qScore = Number(questStatus.todayQuest.score);
          }
        }
      } catch (err) {
        Logger.log("돌발 퀘스트 제목/점수 스캔 중 에러: " + err.toString());
      }
    }
    if (!qTitle && category === "돌발") {
      qTitle = "오아시스 돌발 퀘스트 수행";
    }

    if (category === "돌발") {
      expAwarded = qScore;
      statType = "perf";
      itemTitle = "🔥 돌발: " + qTitle;
    } else if (category === "셀프칭찬") {
      expAwarded = 3; // [v55.0] 나를 사랑하기 3점!
      statType = "def";
      itemTitle = "🌸 셀프 칭찬의 샘 기록";
    } else if (category === "명언" || category === "격언") {
      expAwarded = 10; // [v55.0] 명언 작성자 10점 큰보상!
      statType = "def";
      itemTitle = "📖 명언/격언 공유 공책 작성";
    } else if (category === "릴레이칭찬") {
      expAwarded = 5;
      statType = "def";
      itemTitle = "🤝 릴레이 칭찬 체인 동참 (" + targetMember + "님 지목)";
    }
    
    if (expAwarded > 0) {
      var recordResult = recordActivityLog({
        phone: phone,
        name: payload.author,
        type: (category === "돌발") ? "퀘스트" : "오아시스",
        item: itemTitle,
        score: expAwarded,
        statType: statType
      });
      if (recordResult && recordResult.success) {
        if (category === "돌발") {
          awardMessage = " (돌발 퀘스트 완료 +" + expAwarded + " EXP가 실시간 적립되었습니다!) ⚡";
        } else if (category !== "셀프칭찬") {
          awardMessage = " (오아시스 보상 +" + expAwarded + " EXP가 실시간 적립되었습니다!) 🌟";
        }
      }
    }
    
    var mainMessage = "오아시스 생명샘에 소중한 마음이 안전하게 기록되었습니다!";
    if (category === "셀프칭찬") {
      mainMessage = "스스로를 칭찬하는 긍정 선언은 자존감을 높이고 스트레스를 완화하여 뇌에 긍정적인 활력을 불어넣는 아주 훌륭한 웰니스 요법입니다. 나를 사랑하고 격려하는 따뜻한 한마디를 건넨 당신의 모습이 참으로 멋집니다! 🌸";
    }
    
    return { success: true, message: mainMessage + awardMessage };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * [v55.0] 오아시스 명언/글 공감 리액션 API
 */
function reactOasisPost(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("오아시스_글");
    if (!sheet) return { success: false, error: "오아시스 글 시트가 존재하지 않습니다." };
    
    var rowIdx = Number(payload.rowIdx);
    var clickerPhone = formatPhoneNumber(payload.phone).replace(/[^0-9]/g, "");
    var clickerName = payload.author || "모험가";
    
    if (!rowIdx || rowIdx < 2 || rowIdx > sheet.getLastRow()) {
      return { success: false, error: "유효하지 않은 게시글 번호입니다." };
    }
    
    var postData = sheet.getRange(rowIdx, 1, 1, 10).getValues()[0];
    var authorPhone = formatPhoneNumber(postData[2]).replace(/[^0-9]/g, "");
    var authorName = postData[3];
    var currentHearts = Number(postData[8] || 0);
    var heartPhones = String(postData[9] || "").trim();
    
    // 🛠️ [자가 치유 (Self-healing) 로직] 
    // 기존 데이터에 과학적 지수 표기법(E+31 등)이 들어있어 망가져 있다면,
    // 정상적인 중복 체크가 불가능하므로 빈값으로 정규화하여 자가 치유합니다.
    if (heartPhones.indexOf("E+") > -1 || heartPhones.indexOf("e+") > -1) {
      heartPhones = "";
    }
    
    // 🛡️ [고도화된 중복 방지 엔진]
    // 0이 떼어진 번호, E+ 표기법 등으로 오염되어 매칭이 실패하는 것을 완벽 방어하기 위해
    // 앞자리 '0'을 떼어낸 숫자만 추출해 대조하거나 서브셋 포함 관계를 판별합니다.
    var cleanClicker = clickerPhone.replace(/^0/, ""); // '1012345678'
    var isAlreadyHearted = false;
    
    if (heartPhones) {
      var phoneList = heartPhones.split(",");
      for (var pIdx = 0; pIdx < phoneList.length; pIdx++) {
        var existingPhone = phoneList[pIdx].trim().replace(/[^0-9]/g, "");
        var cleanExisting = existingPhone.replace(/^0/, "");
        
        if (cleanExisting && cleanClicker) {
          // 번호 전체가 정확히 같거나, 앞의 0이 떨어진 형태가 완전히 매칭될 때
          if (cleanExisting === cleanClicker || existingPhone === clickerPhone) {
            isAlreadyHearted = true;
            break;
          }
        }
      }
    }
    
    if (isAlreadyHearted) {
      return { success: false, error: "이미 마음에 스며든 명언입니다. ❤️" };
    }
    
    // 1. 공감클릭자 리스트에 연락처 추가 및 공감수 1 증가
    var newHeartPhones = heartPhones ? (heartPhones + "," + clickerPhone) : clickerPhone;
    var newHearts = currentHearts + 1;
    
    sheet.getRange(rowIdx, 9).setValue(newHearts);
    
    // ✍️ [문자열 포맷 강제] J열에 값을 쓸 때 앞에 `'` 접두사를 붙여 
    // 구글 시트가 대단히 큰 전화번호 문자열 목록을 숫자로 오인하여 E+31 지수 표기로 자동 변환하는 참사를 원천 차단합니다!
    sheet.getRange(rowIdx, 10).setValue("'" + newHeartPhones);
    
    // 2. 리액션을 누른 나에게 +1 EXP 적립
    recordActivityLog({
      phone: clickerPhone,
      name: clickerName,
      type: "오아시스",
      item: "❤️ 명언 공감하기 참여",
      score: 1,
      statType: "def"
    });
    
    // 3. 리액션을 받은 원작자에게 +1 EXP 보너스 적립!
    if (authorPhone && authorPhone !== clickerPhone) {
      recordActivityLog({
        phone: authorPhone,
        name: authorName,
        type: "오아시스",
        item: "✨ 내가 올린 명언에 동료가 공감함",
        score: 1,
        statType: "def"
      });
    }
    
    return { success: true, hearts: newHearts, message: "마음의 온기가 성공적으로 스며들었습니다! (+1 EXP 적립) ❤️" };
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

function addWisdomReaction(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("지혜의_보물고");
    var rowIdx = Number(payload.tipId) + 1; // 1-indexed header + row
    
    var colIdx = payload.type === "like" ? 7 : 8; // G열: 공감수, H열: 깨달음수
    var curVal = Number(sheet.getRange(rowIdx, colIdx).getValue() || 0);
    sheet.getRange(rowIdx, colIdx).setValue(curVal + 1);
    
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function addWisdomComment(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var commentSheet = ss.getSheetByName("지혜의_보물고_댓글") || ss.insertSheet("지혜의_보물고_댓글");
    if (commentSheet.getLastRow() === 0) {
      commentSheet.appendRow(["날짜", "글ID", "작성자", "댓글내용"]);
      commentSheet.setFrozenRows(1);
    }
    
    var now = new Date();
    commentSheet.appendRow([
      Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm:ss"),
      payload.tipId,
      payload.author,
      payload.content
    ]);
    
    return { success: true, message: "칭찬 한마디가 건강 비법서에 기록되었습니다! ✍️" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ==========================================
 * 🛡️ 이장의 특별 축복 실시간 리액션 API
 * ==========================================
 */
function blessActivityReaction(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var archiveSheet = ss.getSheetByName("아카이브");
    var rowIdx = Number(payload.rowIdx);
    
    // 1. 아카이브 시트 리액션 저장 업데이트
    var reactionJson = String(archiveSheet.getRange(rowIdx, 10).getValue() || "{}");
    var reactions = { cool: [], best: [], cheer: [] };
    try { reactions = JSON.parse(reactionJson); } catch(e) {}
    
    if (!reactions[payload.reactionType]) reactions[payload.reactionType] = [];
    
    // 이장님의 특별 도장
    var chiefTag = "CHIEF_" + payload.reactionType;
    if (reactions[payload.reactionType].indexOf(chiefTag) === -1) {
      reactions[payload.reactionType].push(chiefTag);
    }
    
    archiveSheet.getRange(rowIdx, 10).setValue(JSON.stringify(reactions));
    
    // 2. 해당 회원에게 축복 보너스 점수 하사 (+5 / +10 중첩 가능)
    var studentName = payload.name;
    var studentPhone = String(payload.phone).replace(/[^0-9]/g, "");
    var bonusScore = Number(payload.score || 5);
    
    var recordResult = recordActivityLog({
      phone: studentPhone,
      name: studentName,
      type: "방어", // 회복력 [회복]으로 합산
      item: "✨ 이장의 특별 축복(" + payload.reactionType + ")",
      score: bonusScore,
      statType: "def" // 회복력 점수로 즉시 반영
    });
    
    return { success: true, message: studentName + " 모험가에게 특별 축복 점수(+" + bonusScore + " EXP)를 하사하셨습니다! 🛡️✨" };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * ==========================================
 * 📅 돌발 퀘스트 날짜별 캘린더 선세팅 API
 * ==========================================
 */
/**
 * ==========================================
 * 📅 [v66.1] 돌발 퀘스트 날짜별 캘린더 선세팅 API (하이브리드 병합 모드)
 * ==========================================
 */

// [이동 완료] normalizeDateStr 함수는 Common_Utils.gs로 이동되었습니다.

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

/**
 * 📅 [매일 스케줄러] 당일 돌발 퀘스트 자동 활성화 및 아침 선포 알림 발송 (v66.1)
 * 매일 아침(예: 06:00 ~ 08:00 사이) 트리거로 자동 동작하도록 설정합니다.
 * 당일 시행일인 '대기' 상태의 이장 퀘스트를 찾아 '진행'으로 전환하고 전체 알림 쪽지를 발송합니다.
 */
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

function getAllMemberNames() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    if (!regSheet) return [];
    var regData = regSheet.getDataRange().getDisplayValues();
    var regCols = getRegColumnIndices(regSheet);
    var names = [];
    for (var i = 1; i < regData.length; i++) {
      var name = String(regData[i][regCols.name]).trim();
      if (name && name !== "모험가") {
        names.push(name);
      }
    }
    // 중복 제거 및 가나다 정렬
    names = names.filter(function(item, pos) { return names.indexOf(item) === pos; });
    names.sort();
    return names;
  } catch (e) {
    return [];
  }
}
function getMyInbodyHistory(phone) {
  try {
    if (!phone) return { error: "로그인이 필요합니다." };
    var cleanPhone = formatPhoneNumber(phone).replace(/[^0-9]/g, "");
    if (!cleanPhone) return { error: "올바른 연락처 정보가 없습니다." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 회원명단 시트에서 목표체중 로드
    var targetWeight = 0;
    var memberSheet = ss.getSheetByName("회원명단");
    if (memberSheet) {
      var mData = memberSheet.getDataRange().getDisplayValues();
      var mCols = getMemberSheetColumnIndices(memberSheet);
      for (var i = 1; i < mData.length; i++) {
        var mPhone = normalizePhoneDigits(mData[i][mCols.phone]);
        if (mPhone === cleanPhone || (cleanPhone.length >= 8 && mPhone.endsWith(cleanPhone.substring(cleanPhone.length - 8)))) {
          targetWeight = parseWeightSafely(mData[i][mCols.targetWeight]);
          break;
        }
      }
    }
    var sheet = ss.getSheetByName("33챌린지_인바디");
    if (!sheet) return { success: false, error: "'33챌린지_인바디' 시트를 찾을 수 없습니다." };

    var data = sheet.getDataRange().getValues();
    var rawRecords = [];

    // 칼럼 매핑 순서 (기존 및 신규 3단 분기 호환 지원)
    // 0: 측정일 | 1: 회원명 | 2: 연락처 | 3: 체중 | 4: 골격근량 | 5: 체지방률 | ... | 비고/등록일 등
    for (var i = 1; i < data.length; i++) {
      var rawPhone = String(data[i][2] || "").trim();
      var rowPhone = formatPhoneNumber(rawPhone).replace(/[^0-9]/g, "");
      if (rowPhone === cleanPhone) {
        var formattedDate = "";
        try {
          var rawDate = data[i][0];
          if (rawDate instanceof Date) {
            var y = rawDate.getFullYear();
            var m = (rawDate.getMonth() + 1).toString().padStart(2, '0');
            var d = rawDate.getDate().toString().padStart(2, '0');
            formattedDate = y + "-" + m + "-" + d;
          } else {
            formattedDate = String(rawDate).split("T")[0];
          }
        } catch (e) {
          formattedDate = String(data[i][0]);
        }

        // 비고 칼럼: 신규 포맷(index 9 - J열) 또는 구포맷(index 7 - H열) 동적 추출
        var memoVal = "";
        if (data[i].length >= 10) {
          memoVal = String(data[i][9] || "");
        } else {
          memoVal = String(data[i][7] || "");
        }

        rawRecords.push({
          date: formattedDate,
          weight: Number(data[i][3]) || 0,
          muscle: Number(data[i][4]) || 0,
          fat: Number(data[i][5]) || 0,
          memo: memoVal
        });
      }
    }

    // 날짜 순 정렬 (오름차순)하여 3단 스탯 변화를 타임라인으로 정밀 연산
    rawRecords.sort(function(a, b) {
      return new Date(a.date) - new Date(b.date);
    });

    // 인바디 계산 헬퍼 복사
    function localInbodyScoreHelper(first, current, targetWeight, scoreType) {
      if (!first || !current) return 0;
      var score = 0;
      var fW = Number(first.weight) || 0;
      var cW = Number(current.weight) || 0;
      var fM = Number(first.muscle) || 0;
      var cM = Number(current.muscle) || 0;
      var fF = Number(first.fat) || 0;
      var cF = Number(current.fat) || 0;
      
      var diffW = Number((fW - cW).toFixed(2)) || 0;
      var diffM = Number((cM - fM).toFixed(2)) || 0;
      var diffFat = Number((fF - cF).toFixed(1)) || 0;
      
      if (diffW > 0) score += (diffW * 10) * 50;
      if (diffM > 0) score += (diffM * 10) * 200;
      if (diffFat > 0) score += (diffFat * 10) * 100;
      
      var firstFatMass = fW * (fF / 100);
      var currentFatMass = cW * (cF / 100);
      var fatLossRate = 0;
      if (firstFatMass > 0 && !isNaN(firstFatMass)) {
        fatLossRate = ((firstFatMass - currentFatMass) / firstFatMass) * 100;
      }
      if (diffW >= 8.0 || (fatLossRate >= 20.0 && !isNaN(fatLossRate))) {
        score += 1500;
      }
      
      // 🏆 체성분 명품 유지 보너스 판정 엔진 (±0.5kg 정교화)
      if (targetWeight && targetWeight > 0) {
        if (scoreType === "monthly") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        } else if (scoreType === "lifetime") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        }
      }
      
      return isNaN(score) ? 0 : score;
    }

    var finalRecords = [];
    for (var i = 0; i < rawRecords.length; i++) {
      var currentInbody = rawRecords[i];
      var currentDate = new Date(currentInbody.date);

      // (1) 평생 최초 기록 구하기
      var firstEver = rawRecords[0];

      // (2) 이번 주 목요일 구하기
      var day = currentDate.getDay();
      var diffToThu = (day + 3) % 7; 
      var startOfWeek = new Date(currentDate.getFullYear(), currentDate.getMonth(), currentDate.getDate() - diffToThu);
      startOfWeek.setHours(0, 0, 0, 0);

      // (3) 지난주 인바디 구하기: startOfWeek 이전의 가장 최근 기록이되, startOfWeek 기준 7.5일 이내인 것!
      var prevBeforeThisWeek = null;
      for (var j = 0; j < i; j++) {
        var r = rawRecords[j];
        var rDate = new Date(r.date);
        if (rDate < startOfWeek) {
          var timeDiff = startOfWeek.getTime() - rDate.getTime();
          var diffDays = timeDiff / (1000 * 60 * 60 * 24);
          if (diffDays <= 7.5) {
            if (!prevBeforeThisWeek || rDate > new Date(prevBeforeThisWeek.date)) {
              prevBeforeThisWeek = r;
            }
          }
        }
      }

      // (4) 월초 시작일 및 1주차 수요일 마감 마일스톤 구하기
      var year = currentDate.getFullYear();
      var month = currentDate.getMonth();
      var startOfMonth = new Date(year, month, 1, 0, 0, 0, 0);
      var firstDayOfWeek = startOfMonth.getDay();
      var diffToWed = (3 - firstDayOfWeek + 7) % 7;
      var baselineDeadline = new Date(year, month, 1 + diffToWed);
      if (typeof isCenterHoliday === "function" && isCenterHoliday(baselineDeadline)) {
        baselineDeadline.setDate(baselineDeadline.getDate() + 1); // 공휴일 시 목요일 연장
      }
      baselineDeadline.setHours(23, 59, 59, 999);

      // (5) 월초 기준 인바디 구하기: 당월 1일 ~ 1주차 마감 수요일 사이의 첫 기록
      var baselineRecord = null;
      for (var j = 0; j <= i; j++) {
        var r = rawRecords[j];
        var rDate = new Date(r.date);
        if (rDate >= startOfMonth && rDate <= baselineDeadline) {
          if (!baselineRecord || rDate < new Date(baselineRecord.date)) {
            baselineRecord = r;
          }
        }
      }

      // 3단 점수 정교 연산
      var weeklyScore = 0;
      var monthlyScore = 0;
      var totalScore = 0;

      if (firstEver && firstEver !== currentInbody) {
        totalScore = localInbodyScoreHelper(firstEver, currentInbody, targetWeight, 'lifetime');
      }
      if (prevBeforeThisWeek) {
        weeklyScore = localInbodyScoreHelper(prevBeforeThisWeek, currentInbody, targetWeight, 'weekly');
      }
      if (baselineRecord && baselineRecord !== currentInbody) {
        if (isDateInLastWeekMonToWed(currentDate)) {
          monthlyScore = localInbodyScoreHelper(baselineRecord, currentInbody, targetWeight, 'monthly');
        } else {
          monthlyScore = 0;
        }
      }

      finalRecords.push({
        date: currentInbody.date,
        weight: currentInbody.weight,
        muscle: currentInbody.muscle,
        fat: currentInbody.fat,
        weeklyScore: weeklyScore,
        monthlyScore: monthlyScore,
        totalScore: totalScore,
        score: weeklyScore + monthlyScore + totalScore,
        memo: currentInbody.memo
      });
    }

    // 최신 순 정렬 (내림차순)
    finalRecords.sort(function(a, b) {
      return new Date(b.date) - new Date(a.date);
    });

    return { success: true, records: finalRecords, targetWeight: targetWeight };
  } catch (e) {
    return { error: e.toString() };
  }
}

// [v58.1] 갤러리 피드 로딩의 완벽한 구버전 하위호환(가교) 함수 주입
// 옛날 잔재 캐시나 웹앱 구버전에서 getActivePhotos를 쏘더라도 백엔드가 당황하지 않고 척척 응답해 줍니다!
function getActivePhotos(payload) {
  return getArchiveFeed(payload);
}

// [v58.4] 회원의 오늘 상세 모험 로그 및 과거 성취도가 있는 7일 히스토리를 솎아오는 명품 데이터 함수 신규 구현
function getUserWellnessActivityHistory(phone) {
  try {
    if (!phone) return { error: "로그인이 필요합니다." };
    var cleanPhone = String(phone).replace(/[^0-9]/g, "");
    if (!cleanPhone) return { error: "올바른 연락처 정보가 없습니다." };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("일일_활동_기록");
    if (!sheet) return { success: false, error: "'일일_활동_기록' 시트를 찾을 수 없습니다." };

    var data = sheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    
    var todayLogs = [];
    var historyLogs = [];

    // 최신 일지가 보통 아래 행에 있으므로 역순으로 탐색합니다.
    for (var i = data.length - 1; i >= 1; i--) {
      var rowPhone = String(data[i][1]).replace(/[^0-9]/g, "");
      if (rowPhone === cleanPhone) {
        var rawDate = data[i][0];
        var formattedDate = "";
        try {
          if (rawDate instanceof Date) {
            formattedDate = Utilities.formatDate(rawDate, "GMT+9", "yyyy-MM-dd");
          } else {
            formattedDate = String(rawDate).split("T")[0];
          }
        } catch (e) {
          formattedDate = String(rawDate);
        }

        var totalScore = Number(data[i][9]) || 0;
        var rawCompleteStr = String(data[i][8] || "");

        // 1. 오늘 날짜 행일 경우: 완료내역을 파싱하여 개별 리스트로 조립
        if (formattedDate === todayStr) {
          if (rawCompleteStr.trim().length > 0) {
            var items = rawCompleteStr.split(",");
            for (var j = 0; j < items.length; j++) {
              var val = items[j].trim();
              if (val) todayLogs.push(val);
            }
          }
        }

        // 2. 활동이 실제 기록된 의미 있는 최근 7일 히스토리 수집 (최대 7개)
        if (historyLogs.length < 7) {
          if (totalScore > 0 || rawCompleteStr.trim().length > 0) {
            var shortDate = formattedDate;
            if (shortDate.length >= 10) {
              shortDate = shortDate.substring(2, 10); // "2026-05-21" -> "26-05-21" 포맷팅!
            }
            historyLogs.push({
              date: shortDate,
              score: totalScore
            });
          }
        }
      }
    }

    return { 
      success: true, 
      todayLogs: todayLogs, 
      historyLogs: historyLogs 
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * [자동 트리거] 원장님이 스프레드시트에서 수동으로 "등록 현황", "회원명단" 또는 "마을_공지" 시트를 고쳤을 때
 * 구글 서버 캐시를 강제로 비워 즉각 반영되도록 처리합니다.
 */
function onEdit(e) {
  try {
    var sheet = e.range.getSheet();
    var sheetName = sheet.getName();
    if (sheetName === "등록 현황" || sheetName === "등록현황" || sheetName === "회원명단") {
      var cache = CacheService.getScriptCache();
      cache.remove("v45_member_registry");
      console.log("⚡ [onEdit] 시트 수정으로 인한 회원 명단 캐시 초기화 성공!");
    } else if (sheetName === "마을_공지" || sheetName === "마을공지") {
      var cache = CacheService.getScriptCache();
      cache.remove("v58_village_notices");
      cache.remove("v58_pillar_notices");
      console.log("⚡ [onEdit] 시트 수정으로 인한 공지 사항 캐시 초기화 성공!");
    }
  } catch(err) {
    console.error("onEdit 캐시 제거 실패: " + err.toString());
  }
}

// [이동 완료] 우편함 알림 및 공용 알림 수신확인 관리 함수들(checkAndCreatePersonalNotificationSheet, checkAndCreateGlobalNotificationSheet, sendGlobalNotification, sendPersonalNotification, getPersonalNotifications, markNotificationsAsRead, hasSentNotificationToday, isAppActiveDuringAbsence)은 Bridge_Login.gs로 이동되었습니다.

/**
 * [v59.1] 한국인 이름의 특성에 맞춰 성을 지능적으로 제거해주는 친근한 이름 변환기
 */
function getFriendlyName(fullName) {
  if (!fullName) return "회원";
  var name = String(fullName).replace(/\d{4}$/, "").trim();
  
  if (/^[가-힣]+$/.test(name)) {
    var len = name.length;
    if (len === 3) {
      return name.substring(1);
    } else if (len === 4) {
      var doubleSurnames = ['남궁', '황보', '제갈', '사공', '선우', '독고', '동방', '서문'];
      var prefix2 = name.substring(0, 2);
      if (doubleSurnames.indexOf(prefix2) !== -1) {
        return name.substring(2);
      } else {
        return name.substring(1);
      }
    }
  }
  return name;
}

/**
 * [v60.0] 대한민국 공식 법정공휴일 판독기
 * Google Calendar API 연동 및 고정식 공휴일 하이브리드 자동 판정
 */
// [이동 완료] globalPublicHolidayMap, globalFlashHolidayMap, loadPublicHolidaysOnce, loadFlashHolidayMapOnce, isKoreanPublicHoliday, isCenterHoliday, calculateInactiveDays 함수/변수들은 Common_Utils.gs로 이동되었습니다.

/**
 * 🎰 [v66.0] 연-월(Year-Month) 시드 기반의 의사 난수 생성기 (PRNG)
 * 동일 월 내에서는 항상 일관된 셔플 결과를 반환하여 데이터 정합성을 보장합니다.
 */
function SeededRandom(seed) {
  var m = 0x80000000; // 2**31
  var a = 1103515245;
  var c = 12345;
  var state = seed ? seed : Math.floor(Math.random() * (m - 1));
  this.next = function() {
    state = (a * state + c) % m;
    return state / (m - 1);
  };
}

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

/**
 * [v65.80] 주어진 날짜가 해당 월의 마지막 주 월~수요일(목요일 마감 연장선 포함) 범위 내에 해당하는지 판별합니다.
 */
function isDateInLastWeekMonToWed(date) {
  if (!date || isNaN(date.getTime())) return false;
  
  var year = date.getFullYear();
  var month = date.getMonth(); // 0-indexed
  
  // 1. 해당 월의 마지막 날짜 구하기
  var lastDay = new Date(year, month + 1, 0);
  
  // 2. 해당 월의 마지막 수요일 구하기
  var lastWed = new Date(year, month, lastDay.getDate());
  while (lastWed.getDay() !== 3) { // 3은 수요일
    lastWed.setDate(lastWed.getDate() - 1);
  }
  
  // 3. 마지막 수요일이 공휴일/휴무일인 경우 목요일로 연장하는 규칙 반영
  var deadline = new Date(lastWed.getTime());
  var hasHoliday = (typeof isCenterHoliday === "function") && isCenterHoliday(lastWed);
  if (hasHoliday) {
    deadline.setDate(deadline.getDate() + 1); // 목요일로 연장
  }
  deadline.setHours(23, 59, 59, 999);
  
  // 4. 마지막 주 월요일 구하기 (마지막 수요일 기준 2일 전)
  var lastMon = new Date(lastWed.getTime());
  lastMon.setDate(lastWed.getDate() - 2);
  lastMon.setHours(0, 0, 0, 0);
  
  return date >= lastMon && date <= deadline;
}

/**
 * [v67.5] 목요일 시작일 및 주차 구하기 전역 헬퍼 함수
 */
function getThuStartOfWeekLocal(date) {
  var d = new Date(date.getTime());
  var day = d.getDay();
  var diffToThu = (day + 3) % 7; 
  d.setDate(d.getDate() - diffToThu);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getWeekStringLocal(date) {
  var thu = getThuStartOfWeekLocal(date);
  var wed = new Date(thu.getTime() + 6 * 24 * 60 * 60 * 1000);
  var year = wed.getFullYear();
  var month = wed.getMonth() + 1;
  
  var firstDayOfMonth = new Date(year, wed.getMonth(), 1);
  var firstWed = new Date(year, wed.getMonth(), 1);
  var firstDayOfWeek = firstDayOfMonth.getDay();
  var diffToFirstWed = (3 - firstDayOfWeek + 7) % 7;
  firstWed.setDate(1 + diffToFirstWed);
  firstWed.setHours(0, 0, 0, 0);
  
  var weekNum = 1;
  if (wed >= firstWed) {
    weekNum = Math.floor((wed.getDate() - firstWed.getDate()) / 7) + 1;
  } else {
    var prevMonth = new Date(year, wed.getMonth(), 0);
    return getWeekStringLocal(prevMonth);
  }
  return month + "월 " + weekNum + "주";
}



/**
 * [v67.5] 아카이브 시트에서 가장 최신 마감된 주차/월의 전원 순위 리스트를 가볍고 빠르게 로드 (체력/실천/회복 전 스탯 동기화 로드)
 */
function getLatestArchivedRankings(sheetName, isMonthly) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return null;
    
    var data = sheet.getDataRange().getValues();
    // 데이터는 insertRowsAfter(1)로 삽입되므로 두 번째 행(인덱스 1)에 가장 최신 주차/월이 존재함!
    var latestPeriod = String(data[1][0] || "").trim();
    if (!latestPeriod) return null;
    
    var list = [];
    for (var i = 1; i < data.length; i++) {
      var p = String(data[i][0] || "").trim();
      if (p === latestPeriod) {
        var record = {
          rank: Number(data[i][1] || 0),
          name: String(data[i][2] || "").trim(),
          health: Number(data[i][4] || 0), // E열: 체력
          perf: Number(data[i][5] || 0),   // F열: 실천력
          def: Number(data[i][6] || 0),    // G열: 회복력
          score: Number(data[i][7] || 0)   // H열: 주간/월간토탈점수
        };
        if (isMonthly) {
          record.attCount = Number(data[i][9] || 0); // J열: 출석기록
        }
        list.push(record);
      }
    }
    
    // 순위 오름차순 정렬
    list.sort(function(a, b) { return a.rank - b.rank; });
    return { period: latestPeriod, rankings: list };
  } catch (e) {
    Logger.log("getLatestArchivedRankings 에러 (" + sheetName + "): " + e.toString());
    return null;
  }
}

/**
 * [v5.0] 명예의 전당 (Hall of Fame) 데이터 실시간 취합 API
 */
function getHallOfFameData(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var now = new Date();
    
    // 회원명단에서 모든 회원의 목표체중 미리 로드 (1회 조회로 극속 처리!)
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
    
    // 1. 날짜 범위 계산 (KST GMT+9 표준 시간 보정 적용)
    var kstNow = new Date(now.getTime() + 9 * 60 * 60 * 1000);
    var curYear = kstNow.getUTCFullYear();
    var curMonth = kstNow.getUTCMonth(); // 0-11
    var curDate = kstNow.getUTCDate();
    var curDay = kstNow.getUTCDay(); // 0-6
    
    // (1) 주간: 목요일 00:00:00 ~ 수요일 23:59:59 (매주 목요일 00:00 리셋)
    var diffToThu = (curDay + 3) % 7; 
    var startOfWeek = new Date(curYear, curMonth, curDate - diffToThu);
    startOfWeek.setHours(0, 0, 0, 0);
    
    // (2) 월간: 매달 1일 00:00:00 ~ 말일 23:59:59
    var startOfMonth = new Date(curYear, curMonth, 1, 0, 0, 0, 0);
    var endOfMonth = new Date(curYear, curMonth + 1, 0, 23, 59, 59, 999);
    
    // 2. 인바디 기록 미리 로드 및 휴대폰 번호 매핑
    var inbodySheet = ss.getSheetByName("33챌린지_인바디");
    var inData = inbodySheet ? inbodySheet.getDataRange().getValues() : [];
    var userInbodyMap = {};
    for (var k = 1; k < inData.length; k++) {
      var rawPhone = String(inData[k][2] || "").trim();
      var phone = formatPhoneNumber(rawPhone).replace(/[^0-9]/g, "");
      if (!phone) continue;
      
      var iDate = new Date(inData[k][0]);
      if (isNaN(iDate.getTime())) continue;
      
      var record = {
        date: iDate,
        weight: Number(inData[k][3] || 0),
        muscle: Number(inData[k][4] || 0),
        fat: Number(inData[k][5] || 0)
      };
      
      if (!userInbodyMap[phone]) {
        userInbodyMap[phone] = [];
      }
      userInbodyMap[phone].push(record);
    }
    
    // 월초 기준 인바디 마감 기한 계산 (1주차 수요일, 공휴일/휴일 시 목요일 연장)
    var year = now.getFullYear();
    var month = now.getMonth();
    var firstDay = new Date(year, month, 1);
    var firstDayOfWeek = firstDay.getDay(); // 0(Sun) ~ 6(Sat)
    var diffToWed = (3 - firstDayOfWeek + 7) % 7;
    var baselineDeadline = new Date(year, month, 1 + diffToWed);
    var hasWedHoliday = isCenterHoliday(baselineDeadline);
    if (hasWedHoliday) {
      baselineDeadline.setDate(baselineDeadline.getDate() + 1); // 목요일로 연장
    }
    baselineDeadline.setHours(23, 59, 59, 999);
    
    // 3. 일일_활동_기록 로드 및 EXP 집계
    var summarySheet = ss.getSheetByName("일일_활동_기록");
    var summaryData = summarySheet ? summarySheet.getDataRange().getValues() : [];
    
    var userScoresMap = {}; // phone -> { name, weeklyAct, monthlyAct, totalAct, sincereAct }
    
    for (var j = 1; j < summaryData.length; j++) {
      var rawPhone = String(summaryData[j][1] || "").trim();
      var phone = formatPhoneNumber(rawPhone).replace(/[^0-9]/g, "");
      if (!phone) continue;
      
      var name = String(summaryData[j][2] || "모험가").replace(/\d{4}$/, "").trim(); // 실명 마스킹 없음
      var recDateRaw = summaryData[j][0];
      var recDate = (recDateRaw instanceof Date) ? recDateRaw : new Date(recDateRaw);
      if (isNaN(recDate.getTime())) continue;
      
      var rowTotal = Number(summaryData[j][9] || 0); // J열: 웰니스총점
      var perfScore = Number(summaryData[j][5] || 0); // F열: 일반실천_합산
      var defScore = Number(summaryData[j][6] || 0); // G열: 일반회복_합산
      
      if (!userScoresMap[phone]) {
        userScoresMap[phone] = {
          name: name,
          weeklyAct: 0,
          monthlyAct: 0,
          totalAct: 0,
          sincereAct: 0 // 실천 + 회복
        };
      }
      
      var entry = userScoresMap[phone];
      entry.totalAct += rowTotal;
      entry.sincereAct += (perfScore + defScore);
      
      // 주간 합산
      if (recDate >= startOfWeek) {
        entry.weeklyAct += rowTotal;
      }
      // 월간 합산
      if (recDate >= startOfMonth && recDate <= endOfMonth) {
        entry.monthlyAct += rowTotal;
      }
    }
    
    // 인바디 계산 헬퍼 복사
    function localInbodyScoreHelper(first, current, targetWeight, scoreType) {
      if (!first || !current) return 0;
      var score = 0;
      var fW = Number(first.weight) || 0;
      var cW = Number(current.weight) || 0;
      var fM = Number(first.muscle) || 0;
      var cM = Number(current.muscle) || 0;
      var fF = Number(first.fat) || 0;
      var cF = Number(current.fat) || 0;
      
      var diffW = Number((fW - cW).toFixed(2)) || 0;
      var diffM = Number((cM - fM).toFixed(2)) || 0;
      var diffFat = Number((fF - cF).toFixed(1)) || 0;
      
      if (diffW > 0) score += (diffW * 10) * 50;
      if (diffM > 0) score += (diffM * 10) * 200;
      if (diffFat > 0) score += (diffFat * 10) * 100;
      
      var firstFatMass = fW * (fF / 100);
      var currentFatMass = cW * (cF / 100);
      var fatLossRate = 0;
      if (firstFatMass > 0 && !isNaN(firstFatMass)) {
        fatLossRate = ((firstFatMass - currentFatMass) / firstFatMass) * 100;
      }
      if (diffW >= 8.0 || (fatLossRate >= 20.0 && !isNaN(fatLossRate))) {
        score += 1500;
      }
      
      // 🏆 체성분 명품 유지 보너스 판정 엔진 (±0.5kg 정교화)
      if (targetWeight && targetWeight > 0) {
        if (scoreType === "monthly") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        } else if (scoreType === "lifetime") {
          // 결산 시점(최신 기록)이 목표체중 상한선(목표체중 + 0.5kg) 이하로 유지 성공 시 1,000점 지급!
          var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
          if (currentLowerOrEqualTarget) {
            score += 1000;
          }
        }
      }
      
      return isNaN(score) ? 0 : score;
    }
    
    // 각 회원에 대한 인바디 점수 반영
    var finalWeeklyScores = [];
    var finalMonthlyScores = [];
    var finalTotalScores = [];
    
    for (var phone in userScoresMap) {
      var entry = userScoresMap[phone];
      var records = userInbodyMap[phone] || [];
      var targetWeight = userTargetWeightMap[phone] || 0;
      
      // 1) 최초 및 최신 구하기 (전체 누적용)
      var firstEver = null;
      var latestEver = null;
      for (var rIdx = 0; rIdx < records.length; rIdx++) {
        var r = records[rIdx];
        if (!firstEver || r.date < firstEver.date) firstEver = r;
        if (!latestEver || r.date > latestEver.date) latestEver = r;
      }
      var inbodyLifetimeScore = localInbodyScoreHelper(firstEver, latestEver, targetWeight, 'lifetime');
      
      // 2) 주간 인바디 구하기
      var inbodyWeeklyScore = 0;
      if (latestEver && latestEver.date >= startOfWeek) {
        var prevBeforeThisWeek = null;
        for (var rIdx = 0; rIdx < records.length; rIdx++) {
          var r = records[rIdx];
          if (r.date < startOfWeek) {
            if (!prevBeforeThisWeek || r.date > prevBeforeThisWeek.date) {
              prevBeforeThisWeek = r;
            }
          }
        }
        var baseRecord = prevBeforeThisWeek || firstEver;
        inbodyWeeklyScore = localInbodyScoreHelper(baseRecord, latestEver, targetWeight, 'weekly');
      }
      
      // 3) 월간 인바디 구하기 (원장님 피드백 반영 - 월초 기준점 Mon~Wed 필수)
      var inbodyMonthlyScore = 0;
      var baselineRecord = null;
      // 월초 기준 기록 탐색
      for (var rIdx = 0; rIdx < records.length; rIdx++) {
        var r = records[rIdx];
        if (r.date >= startOfMonth && r.date <= baselineDeadline) {
          var rDay = r.date.getDay();
          var isValidDay = (rDay >= 1 && rDay <= 3) || (hasWedHoliday && rDay === 4);
          if (isValidDay) {
            if (!baselineRecord || r.date < baselineRecord.date) {
              baselineRecord = r;
            }
          }
        }
      }
      
      if (baselineRecord) {
        // 당월 말일까지의 최신 인바디 탐색 (엄격 마감 규칙 적용)
        var latestInMonth = null;
        for (var rIdx = 0; rIdx < records.length; rIdx++) {
          var r = records[rIdx];
          if (r.date >= startOfMonth && isDateInLastWeekMonToWed(r.date)) {
            if (!latestInMonth || r.date > latestInMonth.date) {
              latestInMonth = r;
            }
          }
        }
        if (latestInMonth) {
          inbodyMonthlyScore = localInbodyScoreHelper(baselineRecord, latestInMonth, targetWeight, 'monthly');
        }
      }
      
      // 4) 최종 EXP 합산
      var totalExp = entry.totalAct + inbodyLifetimeScore;
      var weeklyExp = entry.weeklyAct + inbodyWeeklyScore;
      var monthlyExp = entry.monthlyAct + inbodyMonthlyScore;
      
      finalWeeklyScores.push({ name: entry.name, score: weeklyExp });
      finalMonthlyScores.push({ name: entry.name, score: monthlyExp });
      finalTotalScores.push({ name: entry.name, score: totalExp });
    }
    
    // 내림차순 정렬 및 순위 부여
    function sortAndRank(arr) {
      arr.sort(function(a, b) { return b.score - a.score; });
      for (var i = 0; i < arr.length; i++) {
        arr[i].rank = i + 1;
      }
      return arr;
    }
    
    sortAndRank(finalWeeklyScores);
    sortAndRank(finalMonthlyScores);
    sortAndRank(finalTotalScores);
    
    // Top 30 필터링
    var topWeekly = finalWeeklyScores.slice(0, 30);
    var topMonthly = finalMonthlyScores.slice(0, 30);
    var topTotal = finalTotalScores; // 토탈은 무제한 노출 스크롤
    
    // [v67.5] 시상식(목요일 / 1일) 당일 아카이브 시트 기반 오버라이딩 비즈니스 로직 (자가치유 실시간 마감 연산 보정 적용)
    var isWeeklyAward = false;
    var weeklyAwardPeriod = "";
    var isMonthlyAward = false;
    var monthlyAwardPeriod = "";
    
    var dayOfWeek = curDay; // 0(일) ~ 6(토)
    var dateOfMonth = curDate; // 1 ~ 31
    
    // [v67.5] 목요일이 평소 발표일이나, 수요일이 휴무였을 경우 금요일로 발표가 연장됩니다.
    var isWeeklyAwardDay = false;
    var awardReferenceDate = null;
    if (dayOfWeek === 4) { // 목요일
      var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
      if (typeof isCenterHoliday !== "function" || !isCenterHoliday(yesterday)) {
        isWeeklyAwardDay = true;
        awardReferenceDate = yesterday;
      }
    } else if (dayOfWeek === 5) { // 금요일
      var dayBeforeYesterday = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000);
      if (typeof isCenterHoliday === "function" && isCenterHoliday(dayBeforeYesterday)) {
        isWeeklyAwardDay = true;
        awardReferenceDate = dayBeforeYesterday;
      }
    }

    // [v67.5] 1일이 평소 월간 발표일이나, 1일이 목요일이고 전날 수요일(말일)이 휴무였을 경우 2일(금요일)로 연장됩니다.
    var isMonthlyAwardDay = false;
    var monthlyAwardReferenceDate = null;
    if (dateOfMonth === 1) { // 1일
      var yesterday = new Date(now.getTime() - 24 * 60 * 60 * 1000);
      if (!(dayOfWeek === 4 && typeof isCenterHoliday === "function" && isCenterHoliday(yesterday))) {
        isMonthlyAwardDay = true;
        monthlyAwardReferenceDate = yesterday;
      }
    } else if (dateOfMonth === 2 && dayOfWeek === 5) { // 2일이 금요일인 경우 (지연된 경우)
      var dayBeforeYesterday = new Date(now.getTime() - 2 * 24 * 60 * 60 * 1000);
      if (typeof isCenterHoliday === "function" && isCenterHoliday(dayBeforeYesterday)) {
        isMonthlyAwardDay = true;
        monthlyAwardReferenceDate = dayBeforeYesterday;
      }
    }
    
    // 이 시점의 아카이브 데이터 임시 보관
    var tempArchiveWeekly = null;
    var tempArchiveMonthly = null;
    
    if (isWeeklyAwardDay && awardReferenceDate) {
      var currentAwardPeriod = getWeekStringLocal(awardReferenceDate); // 예: "5월 4주"
      
      var archivedWeekly = getLatestArchivedRankings("33챌린지_주간성적아카이브", false);
      if (archivedWeekly && archivedWeekly.period === currentAwardPeriod) {
        topWeekly = archivedWeekly.rankings;
        isWeeklyAward = true;
        weeklyAwardPeriod = archivedWeekly.period;
        tempArchiveWeekly = archivedWeekly;
        Logger.log("주간 시상식 오버라이드 (시트 기반): " + weeklyAwardPeriod);
      } else {
        // 아카이브 시트에 오늘 발표할 주차(5월 4주)가 아직 없으면 (트리거 돌기 전), 실시간으로 지난주차 최종 점수 정밀 연산!
        var realTimeWeekly = calculateWeeklyRankingForPeriod(currentAwardPeriod);
        if (realTimeWeekly && realTimeWeekly.rankings.length > 0) {
          topWeekly = realTimeWeekly.rankings;
          isWeeklyAward = true;
          weeklyAwardPeriod = currentAwardPeriod;
          tempArchiveWeekly = realTimeWeekly;
          Logger.log("주간 시상식 오버라이드 (자가치유 실시간 연산 기반): " + weeklyAwardPeriod);
        }
      }
    }
    
    if (isMonthlyAwardDay && monthlyAwardReferenceDate) {
      var currentAwardPeriod = (monthlyAwardReferenceDate.getMonth() + 1) + "월"; // 예: "5월"
      
      var archivedMonthly = getLatestArchivedRankings("33챌린지_월간성적아카이브", true);
      if (archivedMonthly && archivedMonthly.period === currentAwardPeriod) {
        topMonthly = archivedMonthly.rankings;
        isMonthlyAward = true;
        monthlyAwardPeriod = archivedMonthly.period;
        tempArchiveMonthly = archivedMonthly;
        Logger.log("월간 시상식 오버라이드 (시트 기반): " + monthlyAwardPeriod);
      } else {
        // 아카이브 시트에 오늘 발표할 월간(5월)이 아직 없으면, 실시간으로 지난달 최종 점수 정밀 연산!
        var realTimeMonthly = calculateMonthlyRankingForPeriod(currentAwardPeriod);
        if (realTimeMonthly && realTimeMonthly.rankings.length > 0) {
          topMonthly = realTimeMonthly.rankings;
          isMonthlyAward = true;
          monthlyAwardPeriod = currentAwardPeriod;
          tempArchiveMonthly = realTimeMonthly;
          Logger.log("월간 시상식 오버라이드 (자가치유 실시간 연산 기반): " + monthlyAwardPeriod);
        }
      }
    }
    
    // 4. MVP 집계 연산
    // (1) 출석기록 집계 (당월 출석왕 vs 역대 출석의 신)
    var attSheet = ss.getSheetByName("출석기록");
    var attData = attSheet ? attSheet.getDataRange().getValues() : [];
    
    var monthlyAttendanceMap = {}; // phone -> { name, count }
    var totalAttendanceMap = {}; // phone -> { name, count }
    
    if (attSheet && attData.length > 1) {
      var attCols = getAttendanceColumnIndices(attSheet);
      for (var aIdx = 1; aIdx < attData.length; aIdx++) {
        var aPhoneRaw = String(attData[aIdx][attCols.phone] || "").trim();
        var aPhone = formatPhoneNumber(aPhoneRaw).replace(/[^0-9]/g, "");
        if (!aPhone) continue;
        
        var aName = String(attData[aIdx][attCols.name] || "모험가").replace(/\d{4}$/, "").trim();
        var aDateRaw = attData[aIdx][attCols.date];
        var aDate = (aDateRaw instanceof Date) ? aDateRaw : new Date(aDateRaw);
        if (isNaN(aDate.getTime())) continue;
        
        // 역대 누적 카운트
        if (!totalAttendanceMap[aPhone]) {
          totalAttendanceMap[aPhone] = { name: aName, count: 0 };
        }
        totalAttendanceMap[aPhone].count++;
        
        // 당월 카운트
        if (aDate >= startOfMonth && aDate <= endOfMonth) {
          if (!monthlyAttendanceMap[aPhone]) {
            monthlyAttendanceMap[aPhone] = { name: aName, count: 0 };
          }
          monthlyAttendanceMap[aPhone].count++;
        }
      }
    }
    
    // 정렬 후 상위 3명 추출
    function extractTop3(map, valKey) {
      var list = [];
      for (var key in map) {
        list.push(map[key]);
      }
      list.sort(function(a, b) { return b[valKey] - a[valKey]; });
      return list.slice(0, 3);
    }
    
    var mvpMonthlyAttendance = extractTop3(monthlyAttendanceMap, "count");
    var mvpTotalAttendance = extractTop3(totalAttendanceMap, "count");
    
    // (2) 실천왕 (체력을 완전히 제외한 누적 실천+회복 점수 기반)
    var sincereList = [];
    for (var phone in userScoresMap) {
      sincereList.push({ name: userScoresMap[phone].name, score: userScoresMap[phone].sincereAct });
    }
    sincereList.sort(function(a, b) { return b.score - a.score; });
    var mvpSincereKing = sincereList.slice(0, 3);
    
    // (3) 오아시스 천사 (누적 글 작성 수 1~3위)
    var oasisSheet = ss.getSheetByName("오아시스_글");
    var oasisData = oasisSheet ? oasisSheet.getDataRange().getValues() : [];
    var oasisMap = {};
    if (oasisSheet && oasisData.length > 1) {
      for (var oIdx = 1; oIdx < oasisData.length; oIdx++) {
        var oPhoneRaw = String(oasisData[oIdx][2] || "").trim();
        var oPhone = formatPhoneNumber(oPhoneRaw).replace(/[^0-9]/g, "");
        if (!oPhone) continue;
        
        var oName = String(oasisData[oIdx][3] || "모험가").replace(/\d{4}$/, "").trim();
        if (!oasisMap[oPhone]) {
          oasisMap[oPhone] = { name: oName, count: 0 };
        }
        oasisMap[oPhone].count++;
      }
    }
    var mvpOasisAngel = extractTop3(oasisMap, "count");
    
    // 5. 역대 명예의 전당 아카이브 완전 자동 연산 (원장님 피드백 반영 - 100% 무개입 자동화)
    var archiveWeekly = [];
    var archiveMonthly = [];
    var archiveAttendance = [];
    
    try {
      // (1) 주간 성적 아카이브 로드
      var wArcSheet = ss.getSheetByName("33챌린지_주간성적아카이브");
      if (wArcSheet && wArcSheet.getLastRow() > 1) {
        var wArcData = wArcSheet.getDataRange().getValues();
        var wGroups = {}; // period -> [ { name, rank, health, perf, def, score } ]
        var wPeriodsSorted = []; // unique periods in order of appearance
        
        for (var rowIdx = 1; rowIdx < wArcData.length; rowIdx++) {
          var pVal = String(wArcData[rowIdx][0] || "").trim();
          if (!pVal) continue;
          
          if (!wGroups[pVal]) {
            wGroups[pVal] = [];
            wPeriodsSorted.push(pVal);
          }
          wGroups[pVal].push({
            rank: Number(wArcData[rowIdx][1] || 0),
            name: String(wArcData[rowIdx][2] || "").trim(),
            health: Number(wArcData[rowIdx][4] || 0),
            perf: Number(wArcData[rowIdx][5] || 0),
            def: Number(wArcData[rowIdx][6] || 0),
            score: Number(wArcData[rowIdx][7] || 0)
          });
        }
        
        // 각 주차별 Winners 가공
        wPeriodsSorted.forEach(function(periodKey) {
          var users = wGroups[periodKey];
          users.sort(function(a, b) { return a.rank - b.rank; });
          
          var winners = [];
          var limit = Math.min(users.length, 3);
          for (var w = 0; w < limit; w++) {
            var u = users[w];
            winners.push((w + 1) + "." + u.name + "(" + u.score.toLocaleString() + "p)");
          }
          if (winners.length > 0) {
            archiveWeekly.push({
              period: periodKey,
              winners: "🥇 " + winners.join(" | "),
              records: users.slice(0, 10) // 상위 10인의 디테일 레코드 반환
            });
          }
        });
      }
      
      // (2) 월간 및 출석 아카이브 로드
      var mArcSheet = ss.getSheetByName("33챌린지_월간성적아카이브");
      if (mArcSheet && mArcSheet.getLastRow() > 1) {
        var mArcData = mArcSheet.getDataRange().getValues();
        var mGroups = {}; // period -> [ { name, rank, health, perf, def, score, attCount } ]
        var mPeriodsSorted = []; // unique periods in order of appearance
        
        for (var rowIdx = 1; rowIdx < mArcData.length; rowIdx++) {
          var pVal = String(mArcData[rowIdx][0] || "").trim();
          if (!pVal) continue;
          
          if (!mGroups[pVal]) {
            mGroups[pVal] = [];
            mPeriodsSorted.push(pVal);
          }
          mGroups[pVal].push({
            rank: Number(mArcData[rowIdx][1] || 0),
            name: String(mArcData[rowIdx][2] || "").trim(),
            health: Number(mArcData[rowIdx][4] || 0),
            perf: Number(mArcData[rowIdx][5] || 0),
            def: Number(mArcData[rowIdx][6] || 0),
            score: Number(mArcData[rowIdx][7] || 0),
            attCount: Number(mArcData[rowIdx][9] || 0) // J열: 출석기록
          });
        }
        
        // 각 월별 Winners 및 출석왕 가공
        mPeriodsSorted.forEach(function(periodKey) {
          var users = mGroups[periodKey];
          
          // A. 월간 랭킹 가공 (순위 오름차순)
          users.sort(function(a, b) { return a.rank - b.rank; });
          var winners = [];
          var limit = Math.min(users.length, 3);
          for (var w = 0; w < limit; w++) {
            var u = users[w];
            winners.push((w + 1) + "." + u.name + "(" + u.score.toLocaleString() + "p)");
          }
          if (winners.length > 0) {
            archiveMonthly.push({
              period: periodKey,
              winners: "🥇 " + winners.join(" | "),
              records: users.slice(0, 10) // 상위 10인의 디테일 레코드 반환
            });
          }
          
          // B. 출석왕 가공 (출석 횟수 내림차순 정렬)
          var attUsers = [].concat(users);
          attUsers.sort(function(a, b) { return b.attCount - a.attCount; });
          var attWinners = [];
          var attLimit = Math.min(attUsers.length, 3);
          for (var w = 0; w < attLimit; w++) {
            var u = attUsers[w];
            var medal = (w === 0) ? "🥇 1위" : (w === 1 ? "🥈 2위" : "🥉 3위");
            attWinners.push(medal + " " + u.name + "(" + u.attCount + "회)");
          }
          if (attWinners.length > 0) {
            archiveAttendance.push({
              period: periodKey,
              winners: attWinners.join(" | "),
              records: attUsers.slice(0, 3) // 출석왕 TOP 3
            });
          }
        });
      }
      
      // [자가치유 보정] 실시간 연산된 tempArchive가 존재하고 아직 아카이브 리스트에 없으면 역대 MVP 리스트 맨앞에 unshift하여 실시간으로 시상식을 띄워줌!
      if (tempArchiveWeekly && tempArchiveWeekly.rankings && tempArchiveWeekly.rankings.length > 0) {
        var existsW = archiveWeekly.some(function(item) {
          return item.period === tempArchiveWeekly.period;
        });
        if (!existsW) {
          var users = tempArchiveWeekly.rankings;
          var winners = [];
          var limit = Math.min(users.length, 3);
          for (var w = 0; w < limit; w++) {
            winners.push((w + 1) + "." + users[w].name + "(" + users[w].score.toLocaleString() + "p)");
          }
          if (winners.length > 0) {
            archiveWeekly.unshift({
              period: tempArchiveWeekly.period,
              winners: "🥇 " + winners.join(" | "),
              records: users.slice(0, 10)
            });
          }
        }
      }
      
      if (tempArchiveMonthly && tempArchiveMonthly.rankings && tempArchiveMonthly.rankings.length > 0) {
        var existsM = archiveMonthly.some(function(item) {
          return item.period === tempArchiveMonthly.period;
        });
        if (!existsM) {
          var users = tempArchiveMonthly.rankings;
          var winners = [];
          var limit = Math.min(users.length, 3);
          for (var w = 0; w < limit; w++) {
            winners.push((w + 1) + "." + users[w].name + "(" + users[w].score.toLocaleString() + "p)");
          }
          if (winners.length > 0) {
            archiveMonthly.unshift({
              period: tempArchiveMonthly.period,
              winners: "🥇 " + winners.join(" | "),
              records: users.slice(0, 10)
            });
          }
        }
        
        var existsAtt = archiveAttendance.some(function(item) {
          return item.period === tempArchiveMonthly.period;
        });
        if (!existsAtt) {
          var attUsers = [].concat(tempArchiveMonthly.rankings);
          attUsers.sort(function(a, b) { return b.attCount - a.attCount; });
          var attWinners = [];
          var attLimit = Math.min(attUsers.length, 3);
          for (var w = 0; w < attLimit; w++) {
            var u = attUsers[w];
            var medal = (w === 0) ? "🥇 1위" : (w === 1 ? "🥈 2위" : "🥉 3위");
            attWinners.push(medal + " " + u.name + "(" + u.attCount + "회)");
          }
          if (attWinners.length > 0) {
            archiveAttendance.unshift({
              period: tempArchiveMonthly.period,
              winners: attWinners.join(" | "),
              records: attUsers.slice(0, 3)
            });
          }
        }
      }
      
    } catch (arcErr) {
      Logger.log("역대 명예의 전당 아카이브 시트 연동 오류: " + arcErr.toString());
    }

    return {
      success: true,
      data: {
        weekly: topWeekly,
        monthly: topMonthly,
        total: topTotal,
        isWeeklyAward: isWeeklyAward,
        weeklyAwardPeriod: weeklyAwardPeriod,
        isMonthlyAward: isMonthlyAward,
        monthlyAwardPeriod: monthlyAwardPeriod,
        archive: {
          weekly: archiveWeekly,
          monthly: archiveMonthly,
          attendance: archiveAttendance
        },
        mvp: {
          monthlyAttendance: mvpMonthlyAttendance,
          totalAttendance: mvpTotalAttendance,
          sincereKing: mvpSincereKing,
          oasisAngel: mvpOasisAngel
        }
      }
    };
    
  } catch (err) {
    return { success: false, error: "명예의 전당 연산 실패: " + err.toString() };
  }
}

/**
 * 📬 [v60.0] 등록현황 시트에서 '진행중' 상태인 전체 회원의 전화번호 추출
 */
function getActiveUserPhones() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    var phones = [];
    if (regSheet) {
      var data = regSheet.getDataRange().getValues();
      var cols = getRegColumnIndices(regSheet);
      for (var i = 1; i < data.length; i++) {
        var status = String(data[i][cols.status] || "").trim();
        if (status === "진행중" || status === "진행 중") {
          var rawPhone = String(data[i][cols.phone] || "").trim();
          var phone = formatPhoneNumber(rawPhone).replace(/[^0-9]/g, "");
          if (phone && phones.indexOf(phone) === -1) {
            phones.push(phone);
          }
        }
      }
    }
    return phones;
  } catch (e) {
    Logger.log("액티브 회원 추출 실패: " + e.toString());
    return [];
  }
}

/**
 * 📬 [v64.10] 매주 목요일 새벽 0시 5분 실행되는 주간 랭킹 자동 공용 우편 발송 트리거 (단 1행 초고속 적재)
 */
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

/**
 * 📬 [v64.10] 매월 1일 새벽 0시 10분 실행되는 월간 랭킹 및 출석왕 자동 공용 우편 발송 트리거 (단 1행 초고속 적재)
 */
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

/**
 * 📬 [v60.0] 주간 랭킹 시상식 쪽지 내용 실시간 생성 및 미리보기 함수
 */
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

/**
 * 📬 [v60.0] 월간/출석왕 랭킹 시상식 쪽지 내용 실시간 생성 및 미리보기 함수
 */
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

/**
 * 📬 [v66.1] 오늘자 돌발 퀘스트 시상식/공지 쪽지 내용 실시간 생성 및 미리보기 함수
 * 수동 예약이 없을 시 실시간 자동 알고리즘 계산기로 가상 미리보기를 지원하여 무중력 상태를 유지합니다.
 */
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

/**
 * 📬 [v66.1] 매일 새벽 0시 1분 실행되는 오늘자 돌발 퀘스트 자동 선포 및 우편 발송 트리거
 * 비어있는 날짜는 자정에 즉시 실시간 계산하여 시트에 영구 적재(Persistent) 후 전체 쪽지를 선포합니다.
 */
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

/**
 * =========================================================================
 * 📊 [v62.0] 33챌린지 주간/월간 성적 아카이브 DB 구축 및 스냅샷 보존 엔진
 * =========================================================================
 */

/**
 * 1. 특정 주차의 상세 성적 및 결산 시점 누적 토탈 점수를 '33챌린지_주간성적아카이브' 시트에 보존
 */
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

/**
 * 2. 특정 월의 상세 성적 및 결산 시점 누적 토탈 점수를 '33챌린지_월간성적아카이브' 시트에 보존
 */
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

/**
 * 3. 과거 모든 주간 및 월간 성적 데이터를 연산하여 스프레드시트에 통째로 정렬 백필(Migration)
 */
function backfillAllRankingArchives() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var summarySheet = ss.getSheetByName("일일_활동_기록");
    if (!summarySheet) return { success: false, error: "일일_활동_기록 시트가 없습니다." };
    
    var summaryData = summarySheet.getDataRange().getValues();
    var now = new Date();
    
    function getThuStartOfWeekLocal(date) {
      var d = new Date(date.getTime());
      var day = d.getDay();
      var diffToThu = (day + 3) % 7; 
      d.setDate(d.getDate() - diffToThu);
      d.setHours(0, 0, 0, 0);
      return d;
    }
    
    function getWeekStringLocal(date) {
      var thu = getThuStartOfWeekLocal(date);
      var wed = new Date(thu.getTime() + 6 * 24 * 60 * 60 * 1000);
      var year = wed.getFullYear();
      var month = wed.getMonth() + 1;
      
      var firstDayOfMonth = new Date(year, wed.getMonth(), 1);
      var firstWed = new Date(year, wed.getMonth(), 1);
      var firstDayOfWeek = firstDayOfMonth.getDay();
      var diffToFirstWed = (3 - firstDayOfWeek + 7) % 7;
      firstWed.setDate(1 + diffToFirstWed);
      firstWed.setHours(0, 0, 0, 0);
      
      var weekNum = 1;
      if (wed >= firstWed) {
        weekNum = Math.floor((wed.getDate() - firstWed.getDate()) / 7) + 1;
      } else {
        var prevMonth = new Date(year, wed.getMonth(), 0);
        return getWeekStringLocal(prevMonth);
      }
      return month + "월 " + weekNum + "주";
    }

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

/**
 * 4. 관리자용 챌린지 성적 및 스냅샷 조회 API (주간/월간 아카이브 데이터 조회)
 */
function getChallengeArchiveData(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var category = payload.category || "weekly"; // "weekly" or "monthly"
    
    function getThuStartOfWeekLocal(date) {
      var d = new Date(date.getTime());
      var day = d.getDay();
      var diffToThu = (day + 3) % 7; 
      d.setDate(d.getDate() - diffToThu);
      d.setHours(0, 0, 0, 0);
      return d;
    }
    
    function getWeekStringLocal(date) {
      var thu = getThuStartOfWeekLocal(date);
      var wed = new Date(thu.getTime() + 6 * 24 * 60 * 60 * 1000);
      var year = wed.getFullYear();
      var month = wed.getMonth() + 1;
      
      var firstDayOfMonth = new Date(year, wed.getMonth(), 1);
      var firstWed = new Date(year, wed.getMonth(), 1);
      var firstDayOfWeek = firstDayOfMonth.getDay();
      var diffToFirstWed = (3 - firstDayOfWeek + 7) % 7;
      firstWed.setDate(1 + diffToFirstWed);
      firstWed.setHours(0, 0, 0, 0);
      
      var weekNum = 1;
      if (wed >= firstWed) {
        weekNum = Math.floor((wed.getDate() - firstWed.getDate()) / 7) + 1;
      } else {
        var prevMonth = new Date(year, wed.getMonth(), 0);
        return getWeekStringLocal(prevMonth);
      }
      return month + "월 " + weekNum + "주";
    }

    var sheetName = category === "weekly" ? "33챌린지_주간성적아카이브" : "33챌린지_월간성적아카이브";
    var sheet = ss.getSheetByName(sheetName);
    
    // 시트가 아예 없거나 헤더만 있다면 즉시 백필 엔진 실행!
    if (!sheet || sheet.getLastRow() <= 1) {
      backfillAllRankingArchives();
      sheet = ss.getSheetByName(sheetName);
    }
    
    if (!sheet) return { success: true, periods: [], records: [], dateRange: "" };
    
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, periods: [], records: [], dateRange: "" };
    
    // (A) 고유 기간 리스트 추출 (최신 순서가 위에 오도록 역순 루프)
    var periodSet = {};
    var periods = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var rawPeriod = data[i][0];
      var period = "";
      if (rawPeriod instanceof Date) {
        period = (category === "weekly") ? getWeekStringLocal(rawPeriod) : (rawPeriod.getMonth() + 1) + "월";
      } else {
        period = String(rawPeriod || "").trim();
      }
      if (period && !periodSet[period]) {
        periodSet[period] = true;
        periods.push(period);
      }
    }
    
    // (B) 특정 주차/월이 선택된 경우 데이터 필터링
    var targetPeriod = payload.period ? String(payload.period).trim() : "";
    var records = [];
    var dateRange = "";
    
    if (targetPeriod) {
      for (var i = 1; i < data.length; i++) {
        var rawPeriod = data[i][0];
        var rowPeriod = "";
        if (rawPeriod instanceof Date) {
          rowPeriod = (category === "weekly") ? getWeekStringLocal(rawPeriod) : (rawPeriod.getMonth() + 1) + "월";
        } else {
          rowPeriod = String(rawPeriod || "").trim();
        }
        
        if (rowPeriod === targetPeriod) {
          records.push({
            rank: Number(data[i][1] || 0),
            name: String(data[i][2] || "").trim(),
            phone: String(data[i][3] || "").trim(),
            health: Number(data[i][4] || 0),
            perf: Number(data[i][5] || 0),
            def: Number(data[i][6] || 0),
            periodTotal: Number(data[i][7] || 0),
            lifetimeTotal: Number(data[i][8] || 0),
            attendanceCount: data[i][9] !== undefined ? Number(data[i][9] || 0) : 0
          });
        }
      }
      // 순위별 오름차순 정렬
      records.sort(function(a, b) { return a.rank - b.rank; });

      // 동적 날짜 범위 연산
      try {
        if (category === "weekly") {
          var summarySheet = ss.getSheetByName("일일_활동_기록");
          var summaryData = summarySheet ? summarySheet.getDataRange().getValues() : [];
          var matchDates = [];
          for (var j = 1; j < summaryData.length; j++) {
            var recDateRaw = summaryData[j][0];
            var recDate = (recDateRaw instanceof Date) ? recDateRaw : new Date(recDateRaw);
            if (isNaN(recDate.getTime())) continue;
            
            if (getWeekStringLocal(recDate) === targetPeriod) {
              matchDates.push(recDate);
            }
          }
          if (matchDates.length > 0) {
            matchDates.sort(function(a, b) { return a - b; });
            var startD = matchDates[0];
            var endD = matchDates[matchDates.length - 1];
            dateRange = (startD.getMonth() + 1) + "월 " + startD.getDate() + "일 ~ " + 
                        (endD.getMonth() + 1) + "월 " + endD.getDate() + "일";
          }
        } else {
          // 월간
          var targetYear = new Date().getFullYear();
          var monthNum = parseInt(targetPeriod.replace("월", "").trim());
          if (!isNaN(monthNum)) {
            var startOfMonth = new Date(targetYear, monthNum - 1, 1);
            var endOfMonth = new Date(targetYear, monthNum, 0);
            dateRange = (startOfMonth.getMonth() + 1) + "월 1일 ~ " + 
                        (endOfMonth.getMonth() + 1) + "월 " + endOfMonth.getDate() + "일";
          }
        }
      } catch (err) {
        Logger.log("getChallengeArchiveData 날짜연산오류: " + err.toString());
      }
    }
    
    // [v6.0] 과거 월간 조회 시, 해당 과거 월(targetPeriod, 예: "5월")의 진짜 출석왕 TOP 3를 아카이브 시트에서 다이렉트로 추출
    var attendanceMvp = [];
    if (category === "monthly" && targetPeriod) {
      try {
        var sortedRecs = records.concat().sort(function(a, b) {
          return b.attendanceCount - a.attendanceCount;
        });
        attendanceMvp = sortedRecs.slice(0, 3).map(function(r) {
          return { name: r.name, count: r.attendanceCount };
        });
      } catch (attErr) {
        Logger.log("getChallengeArchiveData 아카이브 출석왕 집계 오류: " + attErr.toString());
      }
    }
    
    return { success: true, periods: periods, records: records, dateRange: dateRange, attendanceMvp: attendanceMvp };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * [관리자 전용] 날짜별 판매(수입) 및 지출 기록 역사 요약 목록 조회
 */
function getSalesHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var salesSheet = ss.getSheetByName("판매내역");
    if (!salesSheet) return { error: "판매내역 시트가 없습니다." };
    
    var data = salesSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, logs: [] };
    
    // 지출 대표 5대 대분류
    var expenseCategories = [
      "임대료&관리비",
      "공과금",
      "렌탈료",
      "운영비",
      "기타지출"
    ];
    
    var salesGroup = {};
    for (var i = 1; i < data.length; i++) {
      var dateRaw = data[i][1];
      var dateStr = "";
      if (dateRaw instanceof Date) {
        dateStr = Utilities.formatDate(dateRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        dateStr = String(dateRaw).split("T")[0].trim();
      }
      if (!dateStr || dateStr === "날짜") continue;
      
      var category = String(data[i][2] || "").trim();
      var amount = Number(data[i][5] || 0);
      
      if (!salesGroup[dateStr]) {
        salesGroup[dateStr] = { date: dateStr, incomeTotal: 0, expenseTotal: 0 };
      }
      
      // 지출 카테고리에 포함되면 지출로 가산, 아니면 수입으로 가산
      if (expenseCategories.indexOf(category) !== -1 || category.indexOf("지출") !== -1) {
        salesGroup[dateStr].expenseTotal += amount;
      } else {
        salesGroup[dateStr].incomeTotal += amount;
      }
    }
    
    // 최신 날짜순 정렬
    var salesKeys = Object.keys(salesGroup).sort().reverse();
    var logs = [];
    salesKeys.forEach(function(key) {
      logs.push(salesGroup[key]);
    });
    
    return { success: true, logs: logs };
  } catch (e) {
    return { error: "손익 내역 조회 오류: " + e.toString() };
  }
}

/**
 * 🎫 [v64.60] 회원의 회원권 종류 이름들과 출석 이력을 기반으로 정밀 이용 성향(jumping / therapy / complex / general)을 판별합니다.
 * - 운동만 / 월권: 점핑만 가능 (jumping)
 * - 테라피: 테라피만 가능 (therapy)
 * - 점핑 30회, 점핑 50회: 점핑 & 테라피 둘 다 가능 (complex)
 * - 상호 상반된 회원권을 동시에 여러 개 보유한 경우: 복합 (complex)
 * - 실제 출석 통계가 있는 경우 이용자의 행동 패턴에 맞춰 세분화
 */
function determineMemberClassInfo(membershipNames, attendanceStats) {
  var hasTherapyTicket = false;
  var hasJumpingOnlyTicket = false;
  var hasComplexTicket = false;
  var hasGeneralJumpingTicket = false;
  
  if (membershipNames && Array.isArray(membershipNames)) {
    for (var i = 0; i < membershipNames.length; i++) {
      var mName = String(membershipNames[i] || "").trim();
      if (!mName) continue;
      
      // 1. 점핑 30회, 점핑 50회 등 횟수제 복합 회원권 검출 (둘 다 가능)
      if (mName.indexOf("30회") !== -1 || mName.indexOf("50회") !== -1) {
        hasComplexTicket = true;
      }
      // 2. 테라피 전용 회원권 검출 (테라피, 원적외선, 반신욕 등)
      else if (mName.indexOf("테라피") !== -1 || mName.indexOf("원적외선") !== -1 || mName.indexOf("반신욕") !== -1) {
        hasTherapyTicket = true;
      }
      // 3. 운동만 / 월권 등 점핑 전용 회원권 검출
      else if (mName.indexOf("운동만") !== -1 || mName.indexOf("월권") !== -1) {
        hasJumpingOnlyTicket = true;
      }
      // 4. 일반적인 점핑/운동 회원권 검출
      else if (mName.indexOf("점핑") !== -1 || mName.indexOf("운동") !== -1) {
        hasGeneralJumpingTicket = true;
      }
    }
  }
  
  // 기본 판별 등급
  var regClass = "general";
  
  // 회원권 구성을 통한 성향 분석
  if (hasComplexTicket) {
    regClass = "complex";
  } else if (hasTherapyTicket && (hasJumpingOnlyTicket || hasGeneralJumpingTicket)) {
    // 테라피권과 점핑권을 동시에 둘 다 가진 경우 -> complex
    regClass = "complex";
  } else if (hasTherapyTicket && !hasJumpingOnlyTicket && !hasGeneralJumpingTicket) {
    regClass = "therapy";
  } else if ((hasJumpingOnlyTicket || hasGeneralJumpingTicket) && !hasTherapyTicket) {
    regClass = "jumping";
  }
  
  // 실제 출석 이력(성향) 데이터가 있다면 정밀 보정
  var hasAttendance = false;
  if (attendanceStats) {
    var tCount = attendanceStats.therapy || 0;
    var jCount = attendanceStats.jumping || 0;
    
    if (tCount > 0 || jCount > 0) {
      hasAttendance = true;
      if (tCount > 0 && jCount > 0) {
        // 둘 다 활발히 이용 중인 경우
        regClass = "complex";
      } else if (tCount > 0 && jCount === 0) {
        // 회원권이 복합(complex)이나 일반(general)이더라도 실제 테라피만 이용했다면 테라피에 포커싱
        regClass = "therapy";
      } else if (jCount > 0 && tCount === 0) {
        // 실제 점핑만 이용했다면 점핑에 포커싱
        regClass = "jumping";
      }
    }
  }
  
  return regClass;
}

// ──────────────────────────────────────────────
// V65.50 초고속 예약 흐름 개선용 병합 API
// ──────────────────────────────────────────────

function getMemberIDListWithTicket(v) { 
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var memberSheet = ss.getSheetByName('회원명단');
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    if (!memberSheet) return [];

    var mData = memberSheet.getDataRange().getValues();
    var regData = regSheet ? regSheet.getDataRange().getValues() : [];
    var cols = regSheet ? getRegColumnIndices(regSheet) : null;
    
    var results = [];
    var seen = {}; // 중복 체크용

    // 등록 현황 데이터 미리 번호별 인덱싱하여 조회 속도 O(N) 최적화
    var regMap = {};
    if (regSheet && cols) {
      for (var i = 1; i < regData.length; i++) {
        var rPhone = String(regData[i][cols.phone]).replace(/[^0-9]/g, "");
        var rStatus = String(regData[i][cols.status]).trim();
        if (rStatus === "진행중" || rStatus === "진행 중") {
          if (!regMap[rPhone]) regMap[rPhone] = [];
          regMap[rPhone].push({
            membership: String(regData[i][cols.membership]),
            remainRaw: regData[i][cols.remain]
          });
        }
      }
    }

    for (var i = 1; i < mData.length; i++) {
      var name = String(mData[i][1]).trim();
      var phoneRaw = String(mData[i][2]).trim();
      var phoneOnly = phoneRaw.replace(/[^0-9]/g, "");
      var phone4 = phoneOnly.slice(-4);
      var status = String(mData[i][10]).trim(); // K열 상태
      
      if (status !== "마감" && status !== "정지" && phone4 === v) {
        var key = name + "|" + phone4;
        if (!seen[key]) {
          // 회원권 유효성 일괄 계산
          var hasValidTicket = false;
          var foundTickets = [];
          
          // 1. 등록현황 확인
          var rList = regMap[phoneOnly] || [];
          rList.forEach(function(r) {
            var membership = r.membership;
            var remainRaw = r.remainRaw;
            var remain = parseInt(remainRaw) || 0;
            if (membership.indexOf("테라피") !== -1 || membership.indexOf("점핑") !== -1 || membership.indexOf("회") !== -1 || membership.indexOf("월권") !== -1) {
              if (remain > 0 || String(remainRaw).indexOf("무제한") !== -1) {
                hasValidTicket = true;
                foundTickets.push(membership + "(" + remainRaw + "회)");
              }
            }
          });
          
          // 2. 보너스권 확인 (회원명단 J열 -> Index 9)
          var bonus = parseInt(mData[i][9]) || 0;
          if (bonus > 0) {
            hasValidTicket = true;
            foundTickets.push("보너스권(" + bonus + "회)");
          }

          results.push({
            displayName: name + "(" + phone4 + ")",
            name: name,
            phone: phoneRaw,
            ticketInfo: {
              success: true,
              hasValidTicket: hasValidTicket,
              tickets: foundTickets.join(", ") || "없음"
            }
          }); 
          seen[key] = true;
        }
      }
    }
    return results;
  } catch(e) { return []; }
}

function getTodayTimetableAndRooms(targetDate) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resSheet = ss.getSheetByName('예약DB');
    var holidaySheet = ss.getSheetByName("벙개테라피 및 휴일 설정");
    if (!resSheet) return { error: "시트를 찾을 수 없습니다." };
    
    var targetNum = targetDate.replace(/[^0-9]/g, "");
    var dateObj = new Date(targetDate);
    var dayOfWeek = dateObj.getDay(); 

    var now = new Date();
    var todayNum = Utilities.formatDate(now, "GMT+9", "yyyyMMdd");
    var isToday = (targetNum === todayNum);

    // 기본 시간표 설정 (토요일 3타임, 평일 7타임)
    var baseTimes = (dayOfWeek === 6) ? ["09:00", "09:50", "10:40"] : ["09:00", "09:50", "10:40", "17:00", "17:50", "18:40", "19:30"];
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
            if (hNote !== "") {
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
      return { isHoliday: true, error: holidayMemo || "오늘은 센터 휴무일입니다." };
    }

    if (extraTimes.length > 0) {
      if (dayOfWeek === 0) baseTimes = extraTimes;
      else {
        extraTimes.forEach(function(et) {
          if (baseTimes.indexOf(et) === -1) baseTimes.push(et);
        });
      }
    } else if (dayOfWeek === 0) {
      return { isHoliday: true, error: "일요일은 쉽니다." };
    }

    // 오늘 날짜인 경우, 현재 시간보다 이미 지난 타임은 필터링
    if (isToday) {
      var currentTimeStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
      baseTimes = baseTimes.filter(function(time) {
        return time >= currentTimeStr;
      });
    }

    baseTimes.sort();

    if (baseTimes.length === 0) {
      return { isHoliday: true, error: "오늘은 센터 운영 시간이 없습니다." };
    }

    var resData = resSheet.getDataRange().getDisplayValues().slice(1);
    
    var formatToTwoDigits = function(timeStr) {
      return timeStr.split(':').map(function(v) {
        var clean = v.replace(/[^0-9]/g, "");
        return clean.length === 1 ? "0" + clean : clean;
      }).join(':').substring(0, 5);
    };

    // 각 시간대별 방 상태 맵핑
    var roomStatusMap = {};
    var validTimes = [];

    // 방 리스트
    var defaultRooms = ["1인실", "2인실 A", "2인실 B"];

    baseTimes.forEach(function(t) {
      var targetT = formatToTwoDigits(t);

      if (isToday) {
        var tParts = targetT.split(':');
        var checkTime = new Date(now.getFullYear(), now.getMonth(), now.getDate(), Number(tParts[0]), Number(tParts[1]));
        var limitTime = new Date(now.getTime() + (20 * 60000));
        if (checkTime < limitTime) return; 
      }

      // 예약 리스트 필터링
      var activeReservations = resData.filter(function(row) {
        var rDateNum = row[3].replace(/[^0-9]/g, "");
        var rTimeRaw = String(row[4]).trim();
        if (!rTimeRaw || rTimeRaw === "") return false;
        var rTimeFormatted = formatToTwoDigits(rTimeRaw);
        return rDateNum === targetNum && rTimeFormatted === targetT && row[9] !== "취소";
      });

      // 만약 예약이 3개 이상 꽉 찼으면 이 시간대 제외
      if (activeReservations.length >= 3) return;

      // 예약된 방 추출
      var bookedRooms = activeReservations.map(function(row) {
        return String(row[7]).trim(); // 배정방 기기
      });

      // 방 상태 맵핑 채우기
      var roomStatusList = defaultRooms.map(function(roomName) {
        var isBooked = bookedRooms.indexOf(roomName) !== -1;
        return { name: roomName, isBooked: isBooked };
      });

      roomStatusMap[targetT] = roomStatusList;

      var parts = targetT.split(':');
      var h = Number(parts[0]);
      var icon = (h >= 12) ? "🌙 " : "☀️ ";
      var display = icon + targetT;
      
      validTimes.push({
        timeValue: targetT,
        display: display,
        isFull: false
      });
    });

    return {
      isHoliday: false,
      times: validTimes,
      roomStatusMap: roomStatusMap
    };
  } catch(e) {
    return { error: "서버 오류: " + e.toString() };
  }
}

/**
 * [v67.51] 전역 인바디/유지 점수 계산 헬퍼 함수
 */
function globalInbodyScoreHelper(first, current, targetWeight, scoreType) {
  if (!first || !current) return 0;
  var score = 0;
  var fW = Number(first.weight) || 0;
  var cW = Number(current.weight) || 0;
  var fM = Number(first.muscle) || 0;
  var cM = Number(current.muscle) || 0;
  var fF = Number(first.fat) || 0;
  var cF = Number(current.fat) || 0;
  
  var diffW = Number((fW - cW).toFixed(2)) || 0;
  var diffM = Number((cM - fM).toFixed(2)) || 0;
  var diffFat = Number((fF - cF).toFixed(1)) || 0;
  
  if (diffW > 0) score += (diffW * 10) * 50;
  if (diffM > 0) score += (diffM * 10) * 200;
  if (diffFat > 0) score += (diffFat * 10) * 100;
  
  var firstFatMass = fW * (fF / 100);
  var currentFatMass = cW * (cF / 100);
  var fatLossRate = 0;
  if (firstFatMass > 0 && !isNaN(firstFatMass)) {
    fatLossRate = ((firstFatMass - currentFatMass) / firstFatMass) * 100;
  }
  if (diffW >= 8.0 || (fatLossRate >= 20.0 && !isNaN(fatLossRate))) {
    score += 1500;
  }
  
  // 🏆 체성분 명품 유지 보너스 판정 엔진 (±0.5kg 정교화)
  if (targetWeight && targetWeight > 0) {
    if (scoreType === "monthly") {
      var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
      if (currentLowerOrEqualTarget) {
        score += 1000;
      }
    } else if (scoreType === "lifetime") {
      var currentLowerOrEqualTarget = (cW <= targetWeight + 0.5);
      if (currentLowerOrEqualTarget) {
        score += 1000;
      }
    }
  }
  
  return isNaN(score) ? 0 : score;
}

/**
 * [v67.5] 특정 주차의 랭킹을 마감 기준(어제 수요일 23:59:59)으로 실시간 정밀 집계 (체력 점수에 인바디 점수 및 보너스 합산!)
 */
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
        inbodyWeeklyScore = globalInbodyScoreHelper(baseRecord, latestInWeek, targetWeight, 'weekly');
      }
      
      var inbodyLifetimeScore = globalInbodyScoreHelper(firstEver, latestEver, targetWeight, 'lifetime');
      
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

/**
 * [v67.5] 특정 월의 랭킹을 마감 기준으로 실시간 정밀 집계 (체력 점수에 인바디 점수 및 보너스 합산!)
 */
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
        inbodyMonthlyScore = globalInbodyScoreHelper(baselineRecord, latestInMonth, targetWeight, 'monthly');
      }
      
      var inbodyLifetimeScore = globalInbodyScoreHelper(firstEver, latestEver, targetWeight, 'lifetime');
      
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

/**
 * [v67.51] 과거 잘못 계산된 5월 주차 명칭("5월 3주" -> "5월 4주" 등)을 시트에서 일괄 보정해주는 원클릭 마이그레이션 함수
 * 구글 앱스 스크립트 에디터에서 이 함수를 선택하고 '실행'을 누르면 기존 시트의 데이터가 즉시 올바르게 마이그레이션됩니다.
 */
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

/**
 * 🧹 [v67.5] 수요일(6월 3일) 휴무일에 이미 조기 정산되어 전송된 6월 1주차 주간 랭킹을 롤백하고,
 * 내일(6월 5일 금요일) 새벽에 정상 재정산 및 발표되도록 예약하는 헬퍼 함수
 */
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

/**
 * [v68.1] 백엔드(구글 서버-to-서버)에서 직접 Pollinations 이미지를 다운로드하여 구글 드라이브로 이관하는 함수
 */
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







