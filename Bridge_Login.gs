/**
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * Mobile App Gateway & Bridge Notifications Module
 */

// ──────────────────────────────────────────────
// 1. 모바일 앱 로그인 및 접속 관리
// ──────────────────────────────────────────────

function searchMembersByDigits(payload) {
  try {
    var digits = String(payload.digits || "").trim();
    if (digits.length < 2) return { success: true, members: [] };
    
    // 공용 검색 엔진 활용
    var matched = searchMemberRegistryByDigits(digits);
    var memberMap = {}; 
    
    for (var i = 0; i < matched.length; i++) {
      var m = matched[i];
      // 기존 조건 유지: 진행중이거나 마감 상태인 경우만 로그인 허용
      var status = m.status || "";
      if (status === "진행중" || status === "진행 중" || status.indexOf("마감") !== -1) {
        if (!memberMap[m.phoneClean]) {
          var phoneHint = m.phoneRaw.length > 4 ? "****" + m.phoneRaw.slice(-4) : m.phoneRaw;
          memberMap[m.phoneClean] = { 
            name: m.name, 
            phone: m.phoneClean,
            hint: phoneHint
          };
        }
      }
    }
    
    var matches = Object.values(memberMap);
    return { success: true, members: matches.slice(0, 20) };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

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
    var sheet = ss.getSheetByName("일일_활동_기록");
    if (!sheet) return 0;
    
    var data = sheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    var activeUsers = new Set();
    
    for (var i = 1; i < data.length; i++) {
      var dateRaw = data[i][0];
      var dateStr = (dateRaw instanceof Date) ? Utilities.formatDate(dateRaw, "GMT+9", "yyyy-MM-dd") : String(dateRaw);
      if (dateStr.indexOf(todayStr) > -1) {
        var score = Number(data[i][9]) || 0; // J열: 웰니스총점
        if (score >= 1) {
          activeUsers.add(normalizePhoneDigits(data[i][1])); // B열: 전화번호
        }
      }
    }
    return activeUsers.size;
  } catch(e) { return 0; }
}

function getAnnouncements() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('공지사항');
    var data = sheet.getDataRange().getValues().slice(1);
    return data.map(function(row) { return row[2]; }).filter(Boolean);
  } catch(e) { return ["반갑습니다!"]; }
}

// ──────────────────────────────────────────────
// 2. Suno BGM 링크 해석기
// ──────────────────────────────────────────────

function resolveSunoUrl(url) {
  if (!url) return url;
  url = String(url).trim();
  if (!url) return "";
  
  var urls = url.split(/[,;\n\r]+/);
  var resolvedUrls = [];
  
  for (var i = 0; i < urls.length; i++) {
    var u = urls[i].trim();
    if (!u) continue;
    
    var resolved = resolveSingleSunoUrl(u);
    resolvedUrls.push(resolved);
  }
  
  return resolvedUrls.join(", ");
}

function resolveSingleSunoUrl(url) {
  if (!url) return url;
  url = url.trim();
  
  if (url.indexOf("cdn1.suno.ai") > -1) {
    return url;
  }
  
  var songRegex = /\/song\/([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})/i;
  var match = url.match(songRegex);
  if (match && match[1]) {
    return "https://cdn1.suno.ai/" + match[1] + ".mp3";
  }
  
  var shortRegex = /suno\.com\/s\/([a-zA-Z0-9]+)/i;
  var shortMatch = url.match(shortRegex);
  if (shortMatch && shortMatch[1]) {
    try {
      var response = UrlFetchApp.fetch(url, {
        followRedirects: false,
        muteHttpExceptions: true
      });
      var headers = response.getHeaders();
      var redirectUrl = headers["Location"] || headers["location"];
      if (redirectUrl) {
        var matchRedirect = redirectUrl.match(songRegex);
        if (matchRedirect && matchRedirect[1]) {
          return "https://cdn1.suno.ai/" + matchRedirect[1] + ".mp3";
        }
      }
    } catch (e) {
      Logger.log("Suno 단축 URL 리다이렉트 해석 실패: " + e.toString());
    }
  }
  
  return url;
}

// ──────────────────────────────────────────────
// 3. 우편함 및 개인 알림 관리
// ──────────────────────────────────────────────

function checkAndCreatePersonalNotificationSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("개인알림");
  if (!sheet) {
    sheet = ss.insertSheet("개인알림");
    var headers = ["ID", "전화번호", "이름", "유형", "제목", "내용", "생성시각", "읽은시각"];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setFontWeight("bold")
         .setBackground("#edf2f7");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function checkAndCreateGlobalNotificationSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("전체_알림_목록") || ss.getSheetByName("전체알림");
  if (!sheet) {
    sheet = ss.insertSheet("전체_알림_목록");
    var headers = ["알림ID", "기록시간", "분류", "제목", "내용", "수신확인"];
    sheet.getRange(1, 1, 1, headers.length)
         .setValues([headers])
         .setFontWeight("bold")
         .setBackground("#e2e8f0");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sendGlobalNotification(type, title, content) {
  try {
    var sheet = checkAndCreateGlobalNotificationSheet();
    var now = new Date();
    var createdAtStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm:ss");
    var globalId = "GLOBAL_" + now.getTime() + "_" + Math.floor(Math.random() * 1000);
    
    sheet.appendRow([
      globalId,
      createdAtStr,
      type || "admin",
      title || "전체 공지가 도착했습니다 📢",
      content || "",
      ""
    ]);
    
    var cache = CacheService.getScriptCache();
    cache.remove("v45_member_registry");
    return { success: true, globalId: globalId };
  } catch (e) {
    Logger.log("전체 알림 생성 실패: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function sendPersonalNotification(phone, type, title, content) {
  try {
    var cleanPhone = String(phone || "").replace(/[^0-9]/g, "");
    if (!cleanPhone) return { success: false, error: "전화번호가 누락되었습니다." };
    
    var sheet = checkAndCreatePersonalNotificationSheet();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var name = "모험가";
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    if (regSheet) {
      var regData = regSheet.getDataRange().getDisplayValues();
      var regCols = getRegColumnIndices(regSheet);
      for (var i = 1; i < regData.length; i++) {
        var sheetPhone = String(regData[i][regCols.phone]).replace(/[^0-9]/g, ""); 
        if (sheetPhone === cleanPhone || (cleanPhone.length >= 8 && sheetPhone.endsWith(cleanPhone.substring(cleanPhone.length - 8)))) {
          name = regData[i][regCols.name];
          break;
        }
      }
    }
    
    var now = new Date();
    var createdAtStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm:ss");
    var notiId = "NOTI_" + now.getTime() + "_" + Math.floor(Math.random() * 1000);
    
    sheet.appendRow([
      notiId,
      "'" + cleanPhone,
      name,
      type || "admin",
      title || "쪽지가 도착했습니다 ✉️",
      content || "",
      createdAtStr,
      ""
    ]);
    
    clearUserDashboardCache(cleanPhone);
    return { success: true, notiId: notiId };
  } catch (e) {
    Logger.log("알림 생성 실패: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function getPersonalNotifications(payload) {
  try {
    var rawPhone = String(payload.phone || "");
    var phone = rawPhone.replace(/[^0-9]/g, ""); 
    if (!phone) return { success: false, error: "전화번호가 없습니다." };
    
    var now = new Date();
    var notifications = [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var personalSheet = checkAndCreatePersonalNotificationSheet();
    var personalLastRow = personalSheet.getLastRow();
    if (personalLastRow >= 2) {
      var personalData = personalSheet.getRange(2, 1, personalLastRow - 1, 8).getDisplayValues();
      for (var i = 0; i < personalData.length; i++) {
        var row = personalData[i];
        var rowPhone = String(row[1] || "").replace(/[^0-9]/g, "");
        
        if (rowPhone === phone || (phone.length >= 8 && rowPhone.endsWith(phone.substring(phone.length - 8)))) {
          var notiId = String(row[0] || "");
          var type = String(row[3] || "admin");
          var title = String(row[4] || "");
          var content = String(row[5] || "");
          var createdAt = String(row[6] || "");
          var readAt = String(row[7] || "");
          var isRead = readAt.length > 0;
          
          var keepNotification = true;
          try {
            var createDate = parseDateTimeSafely(createdAt);
            var diffHours = (now.getTime() - createDate.getTime()) / (1000 * 60 * 60);
            
            if (isRead) {
              var readDate = parseDateTimeSafely(readAt);
              var readDiffHours = (now.getTime() - readDate.getTime()) / (1000 * 60 * 60);
              if (readDiffHours > 24) keepNotification = false;
            } else {
              if (diffHours > 72) keepNotification = false;
            }
          } catch (err) {
            keepNotification = true;
          }
          
          if (keepNotification) {
            notifications.push({
              id: notiId,
              type: type,
              title: title,
              content: content,
              createdAt: createdAt,
              isRead: isRead,
              readAt: readAt
            });
          }
        }
      }
    }
    
    var globalSheet = ss.getSheetByName("전체_알림_목록") || ss.getSheetByName("전체알림목록") || ss.getSheetByName("전체알림");
    if (globalSheet) {
      var globalLastRow = globalSheet.getLastRow();
      if (globalLastRow >= 2) {
        var globalData = globalSheet.getRange(2, 1, globalLastRow - 1, 6).getDisplayValues();
        for (var j = 0; j < globalData.length; j++) {
          var gRow = globalData[j];
          var gNotiId = String(gRow[0] || "");
          var gCreatedAt = String(gRow[1] || "");
          var gType = String(gRow[2] || "admin");
          var gTitle = String(gRow[3] || "");
          var gContent = String(gRow[4] || "");
          var gReceivers = String(gRow[5] || ""); 
          
          var gIsRead = gReceivers.indexOf(phone) !== -1;
          var gReadAt = gIsRead ? gCreatedAt : ""; 
          
          var keepGlobal = true;
          try {
            var gDate = parseDateTimeSafely(gCreatedAt);
            var gDiffHours = (now.getTime() - gDate.getTime()) / (1000 * 60 * 60);
            
            if (gIsRead) {
              if (gDiffHours > 24) keepGlobal = false;
            } else {
              if (gDiffHours > 72) keepGlobal = false;
            }
          } catch (err) {
            keepGlobal = true;
          }
          
          if (keepGlobal) {
            notifications.push({
              id: gNotiId,
              type: gType,
              title: gTitle,
              content: gContent,
              createdAt: gCreatedAt,
              isRead: gIsRead,
              readAt: gReadAt
            });
          }
        }
      }
    }
    
    notifications.sort(function(a, b) {
      var dateA = parseDateTimeSafely(a.createdAt);
      var dateB = parseDateTimeSafely(b.createdAt);
      return dateB.getTime() - dateA.getTime();
    });
    
    return { success: true, notifications: notifications };
  } catch (e) {
    Logger.log("알림 목록 조회 실패: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function markNotificationsAsRead(payload) {
  var lock = LockService.getScriptLock();
  try {
    var rawPhone = String(payload.phone || "");
    var phone = rawPhone.replace(/[^0-9]/g, ""); 
    if (!phone) return { success: false, error: "전화번호가 없습니다." };
    
    var now = new Date();
    var nowStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm:ss");
    var updated = false;
    
    lock.waitLock(5000);
    
    var personalSheet = checkAndCreatePersonalNotificationSheet();
    var personalLastRow = personalSheet.getLastRow();
    if (personalLastRow >= 2) {
      var personalRange = personalSheet.getRange(2, 1, personalLastRow - 1, 8);
      var personalData = personalRange.getValues();
      var readAtValues = [];
      
      for (var i = 0; i < personalData.length; i++) {
        var rowPhone = String(personalData[i][1] || "").replace(/[^0-9]/g, "");
        var readAt = String(personalData[i][7] || "");
        
        if ((rowPhone === phone || (phone.length >= 8 && rowPhone.endsWith(phone.substring(phone.length - 8)))) && readAt.length === 0) {
          readAtValues.push([nowStr]);
          updated = true;
        } else {
          var origVal = personalData[i][7];
          if (origVal instanceof Date) {
            origVal = Utilities.formatDate(origVal, "GMT+9", "yyyy-MM-dd HH:mm:ss");
          }
          readAtValues.push([origVal ? String(origVal) : ""]);
        }
      }
      if (updated) {
        personalSheet.getRange(2, 8, personalLastRow - 1, 1).setValues(readAtValues);
      }
    }
    
    var globalSheet = checkAndCreateGlobalNotificationSheet();
    var globalLastRow = globalSheet.getLastRow();
    var globalUpdated = false;
    
    if (globalLastRow >= 2) {
      var globalRange = globalSheet.getRange(2, 1, globalLastRow - 1, 6);
      var globalData = globalRange.getValues();
      var receiverValues = [];
      
      for (var j = 0; j < globalData.length; j++) {
        var receivers = String(globalData[j][5] || ""); 
        
        if (receivers.indexOf(phone) === -1) {
          var newReceivers = receivers + (receivers.length > 0 && !receivers.endsWith("/") ? "/" : "") + phone + "/";
          receiverValues.push([newReceivers]);
          globalUpdated = true;
        } else {
          receiverValues.push([receivers]);
        }
      }
      
      if (globalUpdated) {
        globalSheet.getRange(2, 6, globalLastRow - 1, 1).setValues(receiverValues);
      }
    }
    
    lock.releaseLock();
    
    if (updated || globalUpdated) {
      clearUserDashboardCache(phone);
    }
    
    return { success: true };
  } catch (e) {
    try {
      lock.releaseLock();
    } catch (lockErr) {}
    Logger.log("알림 읽음 처리 실패: " + e.toString());
    return { success: false, error: e.toString() };
  }
}

function hasSentNotificationToday(phone, type, titleKeyword) {
  try {
    var sheet = checkAndCreatePersonalNotificationSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    
    var data = sheet.getRange(2, 1, lastRow - 1, 8).getDisplayValues();
    var cleanPhone = phone.replace(/[^0-9]/g, "");
    var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    
    for (var i = 0; i < data.length; i++) {
      var rowPhone = String(data[i][1] || "").replace(/[^0-9]/g, "");
      var rowType = data[i][3];
      var rowTitle = data[i][4];
      var rowDate = data[i][6].substring(0, 10); 
      
      if (rowPhone === cleanPhone && rowType === type && rowDate === todayStr) {
        if (!titleKeyword || rowTitle.indexOf(titleKeyword) !== -1) {
          return true;
        }
      }
    }
  } catch (e) {
    console.error("Error in hasSentNotificationToday: " + e.toString());
  }
  return false;
}

function isAppActiveDuringAbsence(phone, lastAttendanceDate, now) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("일일_활동_기록");
    if (!sheet) return false;
    
    var data = sheet.getDataRange().getValues();
    var normalizedPhone = phone.replace(/[^0-9]/g, "");
    
    var startDate = new Date(lastAttendanceDate.getTime());
    startDate.setDate(startDate.getDate() + 1);
    startDate.setHours(0, 0, 0, 0);
    
    var endDate = new Date(now.getTime());
    endDate.setHours(23, 59, 59, 999);
    
    for (var i = 1; i < data.length; i++) {
      var rowDateVal = data[i][0];
      if (!rowDateVal) continue;
      
      var rowDate = new Date(rowDateVal);
      rowDate.setHours(0, 0, 0, 0);
      
      if (rowDate >= startDate && rowDate <= endDate) {
        var rowPhone = String(data[i][1] || "").replace(/[^0-9]/g, "");
        if (rowPhone === normalizedPhone) {
          var detail = String(data[i][8] || "");
          var isOnlyOffline = (detail.indexOf("센터방문") !== -1 || detail.indexOf("퇴실") !== -1 || detail.indexOf("테라피") !== -1) && 
                              (detail.indexOf("로그인") === -1 && detail.indexOf("체크") === -1 && detail.indexOf("식단") === -1 && detail.indexOf("인증") === -1 && detail.indexOf("오아시스") === -1 && detail.indexOf("방어") === -1);
          if (!isOnlyOffline) {
            return true; 
          }
        }
      }
    }
  } catch (e) {
    console.error("Error in isAppActiveDuringAbsence: " + e.toString());
  }
  return false;
}


// ──────────────────────────────────────────────
// [이관] getArchiveFeed
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] updateReaction
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] submitArchive
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] getUserDashboardData
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] clearUserDashboardCache
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] submitInBodyRecord
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] get33ChallengeStatus
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] submit33Action
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] getActiveEvents
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] recordActivityLog
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] getOasisPosts
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] submitOasisPost
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] reactOasisPost
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] addWisdomReaction
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] addWisdomComment
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] blessActivityReaction
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] getMyInbodyHistory
// ──────────────────────────────────────────────
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
        totalScore = calculateInbodyScoreHelper(firstEver, currentInbody, targetWeight, 'lifetime');
      }
      if (prevBeforeThisWeek) {
        weeklyScore = calculateInbodyScoreHelper(prevBeforeThisWeek, currentInbody, targetWeight, 'weekly');
      }
      if (baselineRecord && baselineRecord !== currentInbody) {
        if (isDateInLastWeekMonToWed(currentDate)) {
          monthlyScore = calculateInbodyScoreHelper(baselineRecord, currentInbody, targetWeight, 'monthly');
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

// ──────────────────────────────────────────────
// [이관] getUserWellnessActivityHistory
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] getLatestArchivedRankings
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// [이관] getHallOfFameData
// ──────────────────────────────────────────────
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
      var inbodyLifetimeScore = calculateInbodyScoreHelper(firstEver, latestEver, targetWeight, 'lifetime');
      
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
        inbodyWeeklyScore = calculateInbodyScoreHelper(baseRecord, latestEver, targetWeight, 'weekly');
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
          inbodyMonthlyScore = calculateInbodyScoreHelper(baselineRecord, latestInMonth, targetWeight, 'monthly');
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

// ──────────────────────────────────────────────
// [이관] getChallengeArchiveData
// ──────────────────────────────────────────────
function getChallengeArchiveData(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var category = payload.category || "weekly"; // "weekly" or "monthly"
    
    

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

// ──────────────────────────────────────────────
// [이관] getTodayTimetableAndRooms
// ──────────────────────────────────────────────
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

// ──────────────────────────────────────────────
// 🔮 테라피 예약 관리 시스템 (예약 전용 개별 사용자 API)
// ──────────────────────────────────────────────

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
      return "휴무:" + (holidayMemo || "오늘은 센터 휴무일입니다.");
    }

    if (extraTimes.length > 0) {
      if (dayOfWeek === 0) baseTimes = extraTimes;
      else {
        extraTimes.forEach(function(et) {
          if (baseTimes.indexOf(et) === -1) baseTimes.push(et);
        });
      }
    } else if (dayOfWeek === 0) {
      return "휴무:일요일은 쉽니다.";
    }

    if (isToday) {
      var currentTimeStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
      baseTimes = baseTimes.filter(function(time) {
        return time >= currentTimeStr;
      });
    }

    baseTimes.sort();

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
    var status = String(row[9] || "").replace(/s/g, ""); 
    
    var dateParts = String(row[3]).match(/d+/g);
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
// 🔮 테라피 예약 관리 시스템 (예약 전용 개별 사용자 API)
// ──────────────────────────────────────────────

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
      return "휴무:" + (holidayMemo || "오늘은 센터 휴무일입니다.");
    }

    if (extraTimes.length > 0) {
      if (dayOfWeek === 0) baseTimes = extraTimes;
      else {
        extraTimes.forEach(function(et) {
          if (baseTimes.indexOf(et) === -1) baseTimes.push(et);
        });
      }
    } else if (dayOfWeek === 0) {
      return "휴무:일요일은 쉽니다.";
    }

    if (isToday) {
      var currentTimeStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
      baseTimes = baseTimes.filter(function(time) {
        return time >= currentTimeStr;
      });
    }

    baseTimes.sort();

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
    var status = String(row[9] || "").replace(/s/g, ""); 
    
    var dateParts = String(row[3]).match(/d+/g);
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
