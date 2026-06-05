/**
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * Common Utilities & Global Entry Points
 */

// ──────────────────────────────────────────────
// 1. 전역 변수 및 초기 캐싱 설정
// ──────────────────────────────────────────────
var globalPublicHolidayMap = null;
var globalFlashHolidayMap = null;

// ──────────────────────────────────────────────
// 2. 구글 앱스 스크립트 웹앱 진입점 (doGet / doPost / handleRequest)
// ──────────────────────────────────────────────

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
    else if (page === 'halloffame') title = "🏆 [명예의 전당]";
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
    
    // [v44.197] Vercel 수송 데이터 정밀 추적 로그
    if (action === 'submitArchive') {
      var imgLen = (payload && payload.image) ? payload.image.length : 0;
      console.log("[v44.197] 수신 데이터 진단 - Action:", action, "Keys:", Object.keys(payload || {}), "ImageLen:", imgLen);
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

// ──────────────────────────────────────────────
// 3. HTML 템플릿 포함 및 스크립트 정보
// ──────────────────────────────────────────────

/**
 * HTML 파일 내에서 다른 파일을 포함(include) 시키는 함수
 * (CSS, JS를 분리해서 관리하기 위함)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl().split('?')[0];
}

// ──────────────────────────────────────────────
// 4. 지능형 컬럼 매핑 헬퍼 함수
// ──────────────────────────────────────────────

/**
 * [지능형] 회원명단 시트의 열을 동적으로 추적하여 인식하는 지능형 컬럼 매퍼
 */
function getMemberSheetColumnIndices(sheet) {
  var fullData = sheet.getRange(1, 1, 5, Math.min(sheet.getLastColumn(), 30)).getValues();
  var map = { name: 1, phone: 2, targetWeight: 14 }; // defaults
  
  for (var r = 0; r < fullData.length; r++) {
    var hasName = false;
    for (var c = 0; c < fullData[r].length; c++) {
      var title = String(fullData[r][c] || "").trim().replace(/\s/g, "");
      if (title.indexOf("이름") !== -1 || title.indexOf("회원명") !== -1) {
        map.name = c;
        hasName = true;
      }
      if (title.indexOf("휴대폰") !== -1 || title.indexOf("전화번호") !== -1 || title.indexOf("연락처") !== -1) {
        map.phone = c;
      }
      if (title.indexOf("목표체중") !== -1 || title.indexOf("목표몸무게") !== -1) {
        map.targetWeight = c;
      }
    }
    if (hasName) {
      break; 
    }
  }
  return map;
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

// ──────────────────────────────────────────────
// 5. 날짜, 시간, 문자열 포맷팅 및 안전 파싱
// ──────────────────────────────────────────────

/**
 * 앞자리 0 잘림을 방지하고 전화번호를 무결하게 통일하는 안전장치
 */
function formatPhoneSafely(phoneVal) {
  if (!phoneVal) return "";
  var numStr = String(phoneVal).replace(/[^0-9]/g, "");
  // 스프레드시트가 010을 10으로 숫자로 삼켜버린 경우 자동 복구
  if (numStr.length === 10 && numStr.indexOf("10") === 0) {
    numStr = "0" + numStr;
  }
  return numStr;
}

/**
 * 32분 저주 해결 및 8:30, 08:30, 8시30분 오타를 08:30 표준으로 단일화하는 포맷터
 */
function formatTimeSafely(timeVal) {
  if (!timeVal) return "시간미정";
  
  // (A) Date 객체인 경우 (32분 저주 파괴)
  if (timeVal instanceof Date) {
    if (timeVal.getFullYear() < 1905) {
      var hrs = timeVal.getHours();
      var mins = timeVal.getMinutes();
      return (hrs < 10 ? "0" + hrs : hrs) + ":" + (mins < 10 ? "0" + mins : mins);
    }
    return Utilities.formatDate(timeVal, "GMT+9", "HH:mm");
  }
  
  // (B) 스프레드시트 자체 소수 숫자인 경우 환산
  if (typeof timeVal === "number" && timeVal < 1) {
    var totalMins = Math.round(timeVal * 24 * 60);
    var hrs = Math.floor(totalMins / 60);
    var mins = totalMins % 60;
    return (hrs < 10 ? "0" + hrs : hrs) + ":" + (mins < 10 ? "0" + mins : mins);
  }
  
  // (C) 텍스트 오타가 섞여 들어왔을 때 정규식 필터링
  var str = String(timeVal).trim();
  
  // "8:30" 또는 "08:30" 형태
  var matchColon = str.match(/(\d{1,2})\s*:\s*(\d{1,2})/);
  if (matchColon) {
    var hrs = parseInt(matchColon[1], 10);
    var mins = parseInt(matchColon[2], 10);
    return (hrs < 10 ? "0" + hrs : hrs) + ":" + (mins < 10 ? "0" + mins : mins);
  }
  
  // "8시 30분" 형태
  var matchKorean = str.match(/(\d{1,2})\s*시\s*(\d{1,2})/);
  if (matchKorean) {
    var hrs = parseInt(matchKorean[1], 10);
    var mins = parseInt(matchKorean[2], 10);
    return (hrs < 10 ? "0" + hrs : hrs) + ":" + (mins < 10 ? "0" + mins : mins);
  }
  
  return str.indexOf(":") !== -1 ? str.split(":").slice(0, 2).join(":") : str;
}

/**
 * 기상천외한 날짜 형식을 표준 yyyy-MM-dd로 강제 고정하는 헬퍼
 */
function formatDateSafely(dateVal) {
  if (!dateVal) return "";
  if (dateVal instanceof Date) {
    return Utilities.formatDate(dateVal, "GMT+9", "yyyy-MM-dd");
  }
  // 문자열인데 2024.5.11 이나 24.5.11 등으로 작성된 경우 정규화
  var str = String(dateVal).trim();
  var match = str.match(/(\d{2,4})[^\d]+(\d{1,2})[^\d]+(\d{1,2})/);
  if (match) {
    var y = match[1];
    if (y.length === 2) y = "20" + y; // 24 -> 2024 보정
    var m = parseInt(match[2], 10);
    var d = parseInt(match[3], 10);
    return y + "-" + (m < 10 ? "0" + m : m) + "-" + (d < 10 ? "0" + d : d);
  }
  return str;
}

/**
 * ⏰ [v64.40] 기상천외한 타임스탬프/날짜 형식(Date 객체, 한글 점 포맷, 슬래시 포맷 등)을 
 * 무조건 무결한 표준 Date 객체로 강제 리턴하는 초강력 헬퍼
 */
function parseDateTimeSafely(val) {
  if (!val) return new Date();
  if (val instanceof Date) return val;
  
  var str = String(val).trim();
  
  // 1. 구글 스프레드시트에서 가끔 "오후 11:50:00" 처럼 올 때 처리
  var isPM = str.indexOf("오후") !== -1 || str.indexOf("PM") !== -1;
  var isAM = str.indexOf("오전") !== -1 || str.indexOf("AM") !== -1;
  
  // 모든 점(.), 슬래시(/), 오전/오후 텍스트를 정규화하여 숫자 배열만 추출
  var digits = str.replace(/[오전후APM]/g, "").match(/\d+/g);
  if (!digits || digits.length < 3) {
    return new Date(); 
  }
  
  var y = digits[0];
  if (y.length === 2) y = "20" + y; // "26" -> "2026"
  var m = parseInt(digits[1], 10) - 1; // 월은 0부터 시작
  var d = parseInt(digits[2], 10);
  
  var hh = digits.length > 3 ? parseInt(digits[3], 10) : 0;
  var mm = digits.length > 4 ? parseInt(digits[4], 10) : 0;
  var ss = digits.length > 5 ? parseInt(digits[5], 10) : 0;
  
  if (isPM && hh < 12) hh += 12;
  if (isAM && hh === 12) hh = 0;
  
  try {
    return new Date(parseInt(y, 10), m, d, hh, mm, ss);
  } catch (e) {
    return new Date();
  }
}

/**
 * ⚡ [초고속 헬퍼] 루프 내 성능 극대화를 위한 순수 숫자 추출 및 010 자동 보정
 */
function normalizePhoneDigits(phoneStr) {
  if (!phoneStr) return "";
  var phone = String(phoneStr).replace(/[^0-9]/g, "");
  if (phone.startsWith("1") && (phone.length === 10 || phone.length === 9)) {
    phone = "0" + phone;
  }
  return phone;
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
 * ⚖️ [체중 안전 파서] "50kg", "50킬로", "목표 50.5", "50.5" 등의 다양한 텍스트에서 숫자(실수)만 안전하게 추출합니다.
 */
function parseWeightSafely(val) {
  if (val === undefined || val === null || val === "") return 0;
  var str = String(val).trim();
  var match = str.match(/([0-9]+(?:\.[0-9]+)?)/);
  return match ? (Number(match[1]) || 0) : 0;
}

function formatPhoneForSms(phoneStr) {
  return formatPhoneNumber(phoneStr);
}

/**
 * 📅 수동/자동 날짜 문자열 규격화 함수
 * "2026. 6. 1.", "2026/06/01", "2026-06-01" 등 다양한 날짜 포맷을 "yyyy-MM-dd" 표준형으로 통일합니다.
 */
function normalizeDateStr(str) {
  if (!str) return "";
  if (str instanceof Date) {
    return Utilities.formatDate(str, "GMT+9", "yyyy-MM-dd");
  }
  str = String(str).trim();
  if (!str) return "";
  
  // 1. 온점(.), 슬래시(/), 대시(-) 혼용 분절 파싱
  var parts = str.split(/[\.\-\/]/);
  // 끝에 온점이 붙어서 생기는 빈 문자열 요소 등 필터링
  parts = parts.filter(function(p) { return p.trim() !== ""; });
  
  if (parts.length >= 3) {
    var y = parts[0].trim();
    var m = parts[1].trim();
    var d = parts[2].trim();
    
    if (y.length === 4 && !isNaN(y) && !isNaN(m) && !isNaN(d)) {
      var mm = parseInt(m, 10);
      var dd = parseInt(d, 10);
      return y + "-" + (mm < 10 ? "0" + mm : mm) + "-" + (dd < 10 ? "0" + dd : dd);
    }
  }
  
  // 2. yyyyMMdd (8자리 숫자형) 파싱
  if (str.length === 8 && !isNaN(str)) {
    var y = str.substring(0, 4);
    var m = str.substring(4, 6);
    var d = str.substring(6, 8);
    return y + "-" + m + "-" + d;
  }
  
  return str;
}

// ──────────────────────────────────────────────
// 6. 휴일 및 공휴일 관리 (대한민국 공식 API & 센터 맞춤 설정)
// ──────────────────────────────────────────────

function loadPublicHolidaysOnce() {
  if (globalPublicHolidayMap) return;
  globalPublicHolidayMap = {};
  
  try {
    var now = new Date();
    // 오늘부터 앞으로 60일간의 대한민국 공식 공휴일을 단 1회의 Google Calendar API 호출로 일괄 캐싱!
    var future = new Date(now.getTime() + 60 * 24 * 60 * 60 * 1000);
    var cal = CalendarApp.getCalendarById("ko.south_korea#holiday@group.v.calendar.google.com");
    if (cal) {
      var events = cal.getEvents(now, future);
      for (var i = 0; i < events.length; i++) {
        var eventDate = events[i].getStartTime();
        var dateStr = Utilities.formatDate(eventDate, "GMT+9", "yyyy-MM-dd");
        globalPublicHolidayMap[dateStr] = true;
      }
      Logger.log("✅ [성능 최적화] 대한민국 공휴일 일괄 조회 및 캐싱 완료 (공휴일 수: " + events.length + ")");
    }
  } catch (e) {
    Logger.log("⚠️ CalendarApp 일괄 캐싱 실패, 로컬 고정 공휴일 룰로 폴백 작동: " + e.toString());
  }
}

function loadFlashHolidayMapOnce() {
  if (globalFlashHolidayMap) return;
  globalFlashHolidayMap = {};
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var flashSheet = ss.getSheetByName('벙개테라피 및 휴일 설정');
    if (flashSheet) {
      // 벙개테라피/휴일 설정을 단 1회 전체 데이터를 가져와서 메모리 맵으로 고속 로드!
      var flashData = flashSheet.getDataRange().getDisplayValues();
      for (var f = 1; f < flashData.length; f++) {
        var rawDate = flashData[f][0];
        var fDateStr = normalizeDateStr(rawDate);
        if (fDateStr) {
          globalFlashHolidayMap[fDateStr] = String(flashData[f][1]).trim(); // "휴무" 또는 "벙개"
        }
      }
      Logger.log("✅ [성능 최적화] 센터 휴무/벙개 설정 시트 일괄 캐싱 완료!");
    }
  } catch (e) {
    Logger.log("⚠️ loadFlashHolidayMapOnce 시트 로드 오류: " + e.toString());
  }
}

function isKoreanPublicHoliday(date) {
  var dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  
  // 1. 공휴일 캐시 맵이 작동한다면 O(1) 초고속 조회
  loadPublicHolidaysOnce();
  if (globalPublicHolidayMap) {
    return !!globalPublicHolidayMap[dateStr];
  }
  
  // 2. 캐시 실패 시 로컬 Fallback: 고정 법정공휴일 패턴 매칭 (신정, 삼일절, 어린이날, 현충일, 광복절, 개천절, 한글날, 성탄절)
  var mm = date.getMonth() + 1;
  var dd = date.getDate();
  var md = mm + "-" + dd;
  var fixedHolidays = [
    "1-1",   // 신정
    "3-1",   // 삼일절
    "5-5",   // 어린이날
    "6-6",   // 현충일
    "8-15",  // 광복절
    "10-3",  // 개천절
    "10-9",  // 한글날
    "12-25"  // 성탄절
  ];
  return (fixedHolidays.indexOf(md) > -1);
}

/**
 * [v5.0] 센터 공식 휴일 및 공휴일 감지기 (메모리 맵 일괄 조회 캐싱 최적화)
 */
function isCenterHoliday(date) {
  var dayOfWeek = date.getDay();
  if (dayOfWeek === 0) return true; // 일요일 기본 휴무
  
  // 1. 대한민국 법정 공휴일 조회 (캐시 연동)
  if (isKoreanPublicHoliday(date)) return true;
  
  // 2. 센터 설정 휴무 조회 (캐시 연동)
  loadFlashHolidayMapOnce();
  var dateStr = Utilities.formatDate(date, "GMT+9", "yyyy-MM-dd");
  if (globalFlashHolidayMap && globalFlashHolidayMap[dateStr] !== undefined) {
    var type = globalFlashHolidayMap[dateStr];
    if (type === "휴무") {
      return true;
    } else if (type === "벙개") {
      return false; // 벙개는 휴무 제외
    }
  }
  
  return false;
}

/**
 * [v67.6] 두 날짜 사이의 실제 결석 일수를 계산합니다. (일요일 및 센터 휴무일 제외)
 * @param {Date} startDate - 마지막 출석일
 * @param {Date} endDate - 기준일 (보통 오늘)
 * @return {number} - 제외일(일요일, 휴무일)을 뺀 실제 결석 일수
 */
function calculateInactiveDays(startDate, endDate) {
  var s = new Date(startDate.getTime());
  s.setHours(0,0,0,0);
  
  var e = new Date(endDate.getTime());
  e.setHours(0,0,0,0);
  
  var inactiveDays = 0;
  // 시작일 다음날부터 종료일까지 루프
  var current = new Date(s.getTime() + 24 * 60 * 60 * 1000);
  while (current <= e) {
    if (typeof isCenterHoliday === "function") {
      if (!isCenterHoliday(current)) {
        inactiveDays++;
      }
    } else {
      var dayOfWeek = current.getDay();
      if (dayOfWeek !== 0) { // 일요일 제외
        inactiveDays++;
      }
    }
    current.setDate(current.getDate() + 1);
  }
  return inactiveDays;
}
