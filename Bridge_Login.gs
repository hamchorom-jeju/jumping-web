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
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록현황") || ss.getSheetByName("등록 현황");
    if (!regSheet) return { success: false, error: "'등록현황' 시트를 찾을 수 없습니다." };
    
    var data = regSheet.getDataRange().getDisplayValues();
    var cols = getRegColumnIndices(regSheet);
    var memberMap = {}; 
    
    for (var i = 1; i < data.length; i++) {
      var name = String(data[i][cols.name] || "모험가").trim(); 
      var phoneRaw = String(data[i][cols.phone] || "").trim(); 
      var phoneOnlyDigits = phoneRaw.replace(/[^0-9]/g, "");
      var status = String(data[i][cols.status] || "").trim(); 
      
      if ((status === "진행중" || status === "진행 중") && phoneOnlyDigits.slice(-4) === digits) {
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
