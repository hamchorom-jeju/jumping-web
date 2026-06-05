/**
 * ==========================================
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * [Admin.gs] 어드민 백엔드 전용 API 엔진 (Step 3 분리)
 * ==========================================
 */

// 백엔드 메인 안전 제어 밸브 (true: 초고속 신형 로직 / false: 100% 검증 레거시 로직)
var USE_FAST_DASHBOARD = true;

/**
 * 1. 대시보드 & 매출 API
 */

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
    var todayNum = Utilities.formatDate(now, "GMT+9", "yyyydd");
    
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

// [이동 완료] submitSalesRecord 함수는 Common_Utils.gs로 공통화 이관되었습니다.

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

/**
 * 특정 날짜의 매출 가져오기
 */
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
    
    // summary를 배열로 변환
    var summaryArr = [];
    for (var key in summary) {
      summaryArr.push({ category: key, amount: summary[key] });
    }
    
    return { success: true, sales: sales, totalAmount: totalAmount, summary: summaryArr };
  } catch (e) {
    return { error: "매출 조회 오류: " + e.toString() };
  }
}

/**
 * 오늘 매출 가져오기
 */
function getTodaySales() {
  var todayStr = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
  return getSalesByDate(todayStr);
}

/**
 * 매출 내역 수정
 */
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

/**
 * 매출 내역 삭제
 */
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

/**
 * 2. 회원 체크인/퇴실 & 조회 API
 */

/**
 * 관리자 퇴실 처리
 */
function processAdminCheckout(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    var regSheet = ss.getSheetByName("등록 현황");
    
    var cols = getAttendanceColumnIndices(logSheet);
    var rowIdx = data.rowIdx;
    var now = new Date();
    var timeStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
    
    logSheet.getRange(rowIdx, cols.classes + 1, 1, 4).setValues([[data.classes, data.timeLog, "귀가", timeStr]]);
    
    var phoneStr = String(logSheet.getRange(rowIdx, cols.phone + 1).getValue()).replace(/[^0-9]/g, "");
    var memberName = logSheet.getRange(rowIdx, cols.name + 1).getValue();
    
    var timeLogVal = parseFloat(data.timeLog) || 0;
    var reasonText = String(logSheet.getRange(rowIdx, cols.reason + 1).getValue() || "");
    var isCombo = (reasonText.indexOf("복합") !== -1);
    
    recordActivityLog({
      phone: phoneStr,
      name: memberName,
      type: "출석",
      item: "센터방문",
      action: "퇴실",
      score: 20,
      statType: "perf"
    });
    
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

    checkAndAwardWeeklyAttendanceBonus(phoneStr, memberName);

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
        var currentReason = String(logSheet.getRange(rowIdx, cols.reason + 1).getValue());
        if (currentReason.indexOf("(추가차감)") === -1) {
          logSheet.getRange(rowIdx, cols.reason + 1).setValue(currentReason + " (추가차감)");
        }
        logSheet.getRange(rowIdx, cols.memo + 1).setValue("기준시간 초과 (1회 추가 차감됨)");
        
        var remainCount = rData[targetRegIdx-1][rCols.remain];
        var isUnlimited = (remainCount === "(무제한)" || remainCount === "무제한");
        var newVal = isUnlimited ? "무제한" : (Number(remainCount) - 1);
        
        if (!isUnlimited) {
          regSheet.getRange(targetRegIdx, rCols.remain + 1).setValue(newVal);
        }
        
        var currentChange = Number(logSheet.getRange("G" + rowIdx).getValue()) || -1;
        logSheet.getRange("G" + rowIdx).setValue(currentChange - 1); 
        logSheet.getRange("H" + rowIdx).setValue(newVal);
      } else {
        logSheet.getRange(rowIdx, cols.memo + 1).setValue("⚠️회원권 못찾음 - 추가차감 실패 ★★★");
      }
    }
    
    SpreadsheetApp.flush(); 
    
    var resSheet = ss.getSheetByName("예약DB");
    var resData = resSheet.getDataRange().getValues();
    var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var todayNum = todayStr.replace(/[^0-9]/g, "");
    
    for (var k = 1; k < resData.length; k++) {
      var rDateVal = resData[k][3];
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
      
      var rPhone = formatPhoneSafely(resData[k][1]);
      var isPhoneMatch = (rPhone === phoneStr) || (rPhone.length >= 8 && phoneStr.length >= 8 && rPhone.slice(-8) === phoneStr.slice(-8));

      if (rDateNum === todayNum && isPhoneMatch) {
        var status = String(resData[k][9] || "");
        if (status.indexOf("테라피") !== -1 || status.indexOf("예약") !== -1) {
          resSheet.getRange(k + 1, 10).setValue("귀가");
          resSheet.getRange(k + 1, 12).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm:ss"));
          break;
        }
      }
    }
    
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

/**
 * 성함 기반 회원 정보 검색 API (어드민용)
 */
function searchMemberByName(nameStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var searchStr = String(nameStr || "").trim().normalize("NFC").replace(/\s/g, "").toLowerCase();
    if (!searchStr) return { success: true, results: [] };

    var resultsMap = {};
    var sheetNames = ["등록 현황", "등록현황"];
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

// [이동 완료] searchAllMembers 함수는 Registration.gs로 분리 이관되었습니다.

/**
 * 관리자 수동 입실
 */
function manualAdminCheckIn(phoneStr, type) {
  return processAttendance(phoneStr, type, false);
}

/**
 * 자동 퇴실 배치 작업
 */
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
    
    var rData = resSheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      var logDateRaw = data[i][cols.date];
      var logDateStr = (logDateRaw instanceof Date) ? Utilities.formatDate(logDateRaw, "GMT+9", "yyyy-MM-dd") : String(logDateRaw).split(" ")[0];
      var status = String(data[i][cols.status]).trim();
      
      if (logDateStr === todayStr && status === "입실") {
        var rowIdx = i + 1;
        var memberName = String(data[i][cols.name]);
        var reason = String(data[i][cols.reason]);
        var isCombo = reason.indexOf("복합") !== -1;
        var isJumping = reason.indexOf("점핑") !== -1 && !isCombo;
        var isTherapy = reason.indexOf("테라피") !== -1 && !isCombo;

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
          var matchedClass = "";
          var memberPhoneStr = String(data[i][cols.phone]).replace(/[^0-9]/g, "");

          for (var k = 1; k < rData.length; k++) {
            var rDateRaw = rData[k][3];
            var rDateStr = (rDateRaw instanceof Date) ? Utilities.formatDate(rDateRaw, "GMT+9", "yyyy-MM-dd") : String(rDateRaw).split(" ")[0];
            var rPhone = String(rData[k][1]).replace(/[^0-9]/g, "");

            if (rDateStr === todayStr && String(rData[k][2]) === memberName && rPhone === memberPhoneStr) {
              matchedClass = String(rData[k][4]).substring(0, 5);
              if (String(rData[k][9]).indexOf("테라피중") !== -1) {
                resSheet.getRange(k + 1, 10).setValue("완료");
                resSheet.getRange(k + 1, 12).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm:ss") + "[자동]");
              }
              break;
            }
          }

          var autoTimeLog = (isJumping || isCombo) ? "1.0" : "";
          logSheet.getRange(rowIdx, cols.classes + 1).setValue(matchedClass); 
          logSheet.getRange(rowIdx, cols.workoutTime + 1).setValue(autoTimeLog); 
          logSheet.getRange(rowIdx, cols.status + 1).setValue("귀가[자동]"); 
          logSheet.getRange(rowIdx, cols.outTime + 1).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm") + "[자동]"); 
        }
      }
    }
  } catch (e) {
    console.error("자동 퇴실 처리 중 오류: " + e.toString());
  }
}

/**
 * 어드민 퇴실 내역 편집/수정
 */
function editAdminCheckout(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    
    var rowIdx = data.rowIdx;
    var cols = getAttendanceColumnIndices(logSheet);
    var phoneStr = String(logSheet.getRange(rowIdx, cols.phone + 1).getValue());
    var currentExtraText = String(logSheet.getRange(rowIdx, cols.memo + 1).getValue());
    var wasExtraDeducted = currentExtraText.indexOf("(추가차감됨)") !== -1;
    
    logSheet.getRange(rowIdx, cols.classes + 1).setValue(data.classes);
    logSheet.getRange(rowIdx, cols.workoutTime + 1).setValue(data.timeLog);
    
    var regSheet = ss.getSheetByName("등록 현황");
    var rData = regSheet.getDataRange().getValues();
    var rCols = getRegColumnIndices(regSheet);
    
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
        var newVal = isUnlimited ? "무제한" : (Number(curVal) - 1);
        if (!isUnlimited) regSheet.getRange(targetRegIdx, rCols.remain + 1).setValue(newVal);
        
        var currentDeduct = Number(logSheet.getRange("G" + rowIdx).getValue()) || -1;
        logSheet.getRange("G" + rowIdx).setValue(currentDeduct - 1); 
        logSheet.getRange("H" + rowIdx).setValue(newVal);
        
        logSheet.getRange(rowIdx, cols.reason + 1).setValue(currentReason + " (추가차감)");
        logSheet.getRange(rowIdx, cols.memo + 1).setValue("기준시간 초과 (1회 추가 차감됨)");
      } 
      else if (wasExtraDeducted && !data.extraDeduct) {
        var newVal = isUnlimited ? "무제한" : (Number(curVal) + 1);
        if (!isUnlimited) regSheet.getRange(targetRegIdx, rCols.remain + 1).setValue(newVal);
        
        var currentDeduct = Number(logSheet.getRange("G" + rowIdx).getValue()) || -2;
        logSheet.getRange("G" + rowIdx).setValue(currentDeduct + 1); 
        logSheet.getRange("H" + rowIdx).setValue(newVal);
        
        logSheet.getRange(rowIdx, cols.reason + 1).setValue(currentReason);
        logSheet.getRange(rowIdx, cols.memo + 1).setValue("");
      }
    } else if (data.extraDeduct) {
      logSheet.getRange(rowIdx, cols.memo + 1).setValue("⚠️회원권 못찾음 - 추가차감 실패 ★★★");
    }
    
    SpreadsheetApp.flush(); 
    
    var currentStatus = String(logSheet.getRange(rowIdx, cols.status + 1).getValue());
    if (currentStatus === "입실") {
      logSheet.getRange(rowIdx, 12).setValue("귀가");
      var now = new Date();
      var timeStr = Utilities.formatDate(now, "GMT+9", "HH:mm");
      logSheet.getRange(rowIdx, 13).setValue(timeStr);
      
      try {
        var todayStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
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

/**
 * 특정 일자의 타임별 출석 회원을 가져옵니다.
 */
function getClassMembersByDate(dateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName("출석기록");
    if (!logSheet) return { error: "출석기록 시트가 없습니다." };

    var data = logSheet.getDataRange().getDisplayValues();
    var cols = getAttendanceColumnIndices(logSheet);
    
    var classMembers = { "09시": [], "10시": [], "17시": [], "18시": [], "19시": [], "20시": [] };
    var allLogs = [];

    for (var i = 1; i < data.length; i++) {
      var logDateStr = data[i][cols.date].split(" ")[0];

      if (logDateStr === dateStr) {
        var mName = String(data[i][cols.name]);
        var mInTime = String(data[i][cols.inTime] || data[i][1] || ""); 
        var mClasses = String(data[i][cols.classes] || "");
        var mType = String(data[i][cols.type] || "");
        var mReason = String(data[i][cols.reason] || "");
        var mTimeLog = String(data[i][cols.workoutTime] || "");
        var mStatus = String(data[i][cols.status] || "");
        
        for (var slot in classMembers) {
          if (mClasses.indexOf(slot) !== -1) {
            classMembers[slot].push(mName);
          }
        }

        allLogs.push({
          name: mName,
          inTime: mInTime,
          membership: mType, 
          type: mType,       
          classes: mClasses, 
          timeLog: mTimeLog, 
          change: String(data[i][cols.change] || "0").replace("-", ""), 
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
 * 3. 회원권 등록 및 리뉴얼 API
 */

// [이동 완료] getMemberRenewalData 함수는 Common_Utils.gs로 공통화 이관되었습니다.

// [이동 완료] submitRegistration 함수는 Registration.gs로 분리 이관되었습니다.

// [이동 완료] generateSmsContent 함수는 Common_Utils.gs로 공통화 이관되었습니다.

/**
 * [연장 전용] 회원권 연장 / 재등록 처리
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
    
    var cData = configSheet.getDataRange().getValues();
    var passInfo = { baseCount: 0, validDays: 0, price: 0 };
    for (var c = 1; c < cData.length; c++) {
      if (cData[c][0] === data.membership) {
        passInfo.baseCount = Number(cData[c][1]) || 0;
        passInfo.validDays = Number(cData[c][2]) || 0; 
        passInfo.price = Number(cData[c][6]) || 0;
        break;
      }
    }

    var regData = regSheet.getDataRange().getValues();
    var targetRowIdx = -1;
    var currentExp = startDateStr;
    var currentRemain = 0;
    var memberName = "";

    for (var k = regData.length - 1; k >= 1; k--) {
      if (String(regData[k][2]) === data.phone) {
        if (!memberName) memberName = regData[k][1];
        
        var isOngoing = (String(regData[k][8]).trim() === "진행중");
        var oldType = String(regData[k][4]).trim();
        var newType = data.membership.trim();
        
        var isSameFamily = (oldType === newType); 
        
        if (!isSameFamily && oldType.indexOf("점핑") !== -1 && oldType.indexOf("회") !== -1 && newType.indexOf("점핑") !== -1 && newType.indexOf("회") !== -1) {
          isSameFamily = true;
        }
        
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

    var calcExpDate = function(baseDateStr, daysToAdd) {
      var d = new Date(baseDateStr);
      d.setDate(d.getDate() + Number(daysToAdd));
      return Utilities.formatDate(d, "GMT+9", "yyyy-MM-dd");
    };

    var finalExp = "";
    var finalRemain = passInfo.baseCount;
    var finalMemo = data.memo || "";
    
    var gapDays = 0;
    if (currentExp && currentExp !== "undefined" && currentExp !== "-") {
      var prevExp = new Date(currentExp);
      gapDays = Math.floor((now.getTime() - prevExp.getTime()) / (1000 * 60 * 60 * 24));
    }

    var curRemainVal = Number(currentRemain) || 0;

    if (data.renewType === "연장결제") {
      var expD = new Date(currentExp);
      expD.setHours(23, 59, 59, 999);
      var baseDate = (now <= expD) ? currentExp : startDateStr;
      finalExp = calcExpDate(baseDate, passInfo.validDays);
      
      var isMonthly = (data.membership.indexOf("월권") !== -1 || data.membership.indexOf("운동만") !== -1);
      
      if (isMonthly) {
        finalRemain = passInfo.baseCount;
        var isNotExpired = (now <= expD);
        
        if (curRemainVal > 0 && isNotExpired) {
          var carryTag = "[월권이월:" + curRemainVal + "회|기한:" + currentExp + "]";
          finalMemo = "잔여 " + curRemainVal + "회 이월딱지 생성 (" + currentExp + "까지)";
          
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
        finalRemain = curRemainVal + passInfo.baseCount;
        finalMemo = "잔여 " + curRemainVal + "회 합산 연장 (" + finalRemain + "회)";
      }
      
      if (targetRowIdx !== -1) {
        var oldMemo = String(regSheet.getRange(targetRowIdx, 11).getValue() || "");
        regSheet.getRange(targetRowIdx, 9).setValue("만료(연장)"); 
        regSheet.getRange(targetRowIdx, 8).setValue(0);
        regSheet.getRange(targetRowIdx, 11).setValue(oldMemo + " / 잔여횟수 " + curRemainVal + "회 연장결제로 인해 차기이월로 0회 처리");
      }
    } else if (data.renewType === "재결제") {
      finalExp = calcExpDate(startDateStr, passInfo.validDays);
      finalRemain = passInfo.baseCount;
      finalMemo = "기존 " + curRemainVal + "회 소멸 후 재결제";
      if (targetRowIdx !== -1) {
        regSheet.getRange(targetRowIdx, 9).setValue("마감(재결제)");
      }
    } else if (data.renewType === "추가결제") {
      finalExp = calcExpDate(startDateStr, passInfo.validDays);
      finalRemain = passInfo.baseCount;
      finalMemo = "추가 구매 (별도 기간/횟수 적용)";
    } else {
      finalExp = calcExpDate(startDateStr, passInfo.validDays);
      finalRemain = passInfo.baseCount;
      finalMemo = "기타 결제";
    }

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
    
    regSheet.appendRow([
      startDateStr,       
      memberName,         
      data.phone,          
      data.renewType,      
      data.membership,     
      startDateStr,       
      finalExp,           
      finalRemain,        
      "진행중",            
      passInfo.price,     
      finalMemo,           
      data.signature       
    ]);

    submitSalesRecord({
      date: startDateStr,
      category: "회원등록",
      buyer: memberName,
      itemName: data.membership,
      amount: data.price,
      payMethod: data.payMethod || "카드",
      memo: (data.renewType === "연장결제" || data.renewType === "재결제") ? "재등록" : "추가등록"
    });

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

/**
 * 4. 업무일지 관리 API
 */

/**
 * 업무일지 데이터 저장/수정
 */
function submitWorkLog(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var workLogSheet = ss.getSheetByName("업무일지");
    var dateStr = data.date || Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    
    var wlData = workLogSheet.getDataRange().getValues();
    var targetRowIdx = -1;
    
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
      workLogSheet.getRange(targetRowIdx, 1, 1, rowValues.length).setValues([rowValues]);
      return { success: true, message: dateStr + " 업무일지가 업데이트되었습니다." };
    } else {
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

/**
 * 업무일지 최신 기록 조회
 */
function getWorkLogHistory() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var workLogSheet = ss.getSheetByName("업무일지");
    if (!workLogSheet) return { error: "업무일지 시트가 없습니다." };
    
    var data = workLogSheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, logs: [] };
    
    var logs = [];
    var start = Math.max(1, data.length - 15);
    for (var i = data.length - 1; i >= start; i--) {
      var row = data[i];
      
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

/**
 * 일일 정산 기준 업무일지 통계 및 기존 일지 검색
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
        stats.total++; 

        var classes = String(data[i][cols.classes] || "");
        var type = String(data[i][cols.type] || "");
        var reason = String(data[i][cols.reason] || "");
        
        for (var slot in stats) {
          if (slot.indexOf("시") !== -1 && classes.indexOf(slot) !== -1) {
            stats[slot]++;
          }
        }

        var isCombo = (type === "복합" || reason.indexOf("복합") !== -1);
        var isTherapy = (type.indexOf("테라피") !== -1 || reason.indexOf("테라피") !== -1 || reason.indexOf("보너스") !== -1) && !isCombo;

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
 * 밤 11시 30분경 자동으로 실행되어 그날의 최종 통계를 업무일지에 업데이트
 */
function autoCloseDailyLog() {
  try {
    var now = new Date();
    var dateStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var dayOfWeek = now.getDay(); 
    var isOff = (dayOfWeek === 0); 
    
    if (!isOff && isKoreanPublicHoliday(now)) {
      isOff = true;
    }
    
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
              isOff = false; 
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
    
    var statsRes = getTodayWorkLogStats(dateStr);
    if (statsRes.error) {
      Logger.log("자동 마감 통계 조회 실패: " + statsRes.error);
      return;
    }
    var stats = statsRes.stats;
    
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
    
    var data = {
      date: dateStr,
      author: existingLog ? String(existingLog[1] || "") : "시스템 자동마감",
      jumpingList: existingLog ? String(existingLog[2] || "") : "(미기입)",
      muscleList: existingLog ? String(existingLog[3] || "") : "(미기입)",
      stats: stats,
      remarks: existingLog ? String(existingLog[13] || "") : "퇴근 전 미작성되어 시스템에 의해 자동 마감되었습니다.",
      issues: existingLog ? String(existingLog[14] || "") : ""
    };
    
    var res = submitWorkLog(data);
    Logger.log(dateStr + " 자동 마감 결과: " + res.message);
    return res;
    
  } catch (e) {
    Logger.log("자동 마감 치명적 오류: " + e.toString());
    return { error: e.toString() };
  }
}

/**
 * 5. 인바디 관리 API
 */

/**
 * 33챌린지 신규 인바디 데이터 등록
 */
function saveInBodyRecord(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("인바디 입력");
    if (!sheet) return { error: "'인바디 입력' 시트가 없습니다. 시트명을 확인해주세요." };
    
    sheet.appendRow([
      "", 
      data.date || Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd"), 
      data.name, 
      data.phone, 
      data.weight, 
      data.muscle, 
      data.fat, 
      data.memo || "" 
    ]);
    
    return { success: true, message: data.name + " 님의 인바디 기록이 저장되었습니다." };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * 33챌린지 인바디 이력 조회
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
      var dateVal = row[0]; 
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, "GMT+9", "yyyy-MM-dd");
      }
      
      var wScore = row[6] !== "" ? Number(row[6]) : 0; 
      var mScore = row[7] !== "" ? Number(row[7]) : 0; 
      var tScore = row[8] !== "" ? Number(row[8]) : 0; 
      
      records.push({
        rowIdx: i + 1,
        date: String(dateVal || ""),
        name: String(row[1] || ""), 
        phone: String(row[2] || ""), 
        weight: row[3] !== "" ? Number(row[3]) : 0, 
        muscle: row[4] !== "" ? Number(row[4]) : 0, 
        fat: row[5] !== "" ? Number(row[5]) : 0, 
        weeklyScore: wScore,
        monthlyScore: mScore,
        totalScore: tScore,
        score: wScore + mScore + tScore, 
        memo: String(row[9] || "") 
      });
    }
    records.reverse();
    return { success: true, records: records };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * 33챌린지 인바디 기록 개별 수정 및 점수 자동 재계산
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
    
    var data = sheet.getDataRange().getValues();
    var userRecords = [];
    
    for (var i = 1; i < data.length; i++) {
      var rowPhone = formatPhoneNumber(data[i][2]).replace(/[^0-9]/g, "");
      if (rowPhone === phone) {
        var rDate = new Date(data[i][0]);
        if (isNaN(rDate.getTime())) continue;
        
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

    userRecords.sort(function(a, b) { return a.date - b.date; });

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
      
      if (targetWeight && targetWeight > 0) {
        if (scoreType === "monthly" || scoreType === "lifetime") {
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
    
    clearUserDashboardCache(phone);
    
    return { success: true, message: "기록이 수정되었으며 점수가 자동 재계산되었습니다." };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}



/**
 * 7. 테라피 예약 & 벙개/휴일 설정 API
 */

/**
 * 특정 날짜의 관리자용 테라피 예약 내역 조회
 */
function getAdminReservationData(targetDateStr) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("예약DB");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    
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
      var rDateStr = data[i][idx.date].split(" ")[0]; 
      
      if (rDateStr === targetDateStr) {
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
          time: timeStr, 
          name: data[i][idx.name],
          room: data[i][idx.room],
          people: data[i][idx.people],
          status: data[i][idx.status]
        });
      }
    }
    
    results.sort((a, b) => {
      if (a.time !== b.time) return a.time.localeCompare(b.time);
      return a.room.localeCompare(b.room);
    });
    
    return results;
  } catch (e) {
    return [];
  }
}

/**
 * 예약 행 삭제
 */
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
 * 관리자용 특정 기간 범위 테라피 예약 조회 범위 API
 */
function getAdminReservationDataRange(startDateStr, durationDays) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("예약DB");
    if (!sheet) return [];
    
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    
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
      var rDateStr = String(data[i][idx.date] || "").split(" ")[0]; 
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
 * 테라피 예약 정보 개별 수정 API
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

// [이동 완료] getFlashSettingByDate, saveFlashSetting, getAllFlashSettings, deleteFlashSettingByRow 함수들은 Common_Utils.gs로 공통화 이관되었습니다.

/**
 * 8. 문자관리 (SMS) API
 */

/**
 * 문자발송 대기 목록 가져오기 (관리자 전용)
 */
function getSmsLogData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("문자발송");
    if (!sheet) return { error: "문자발송 시트가 없습니다." };
    
    var data = sheet.getDataRange().getDisplayValues();
    var results = [];
    
    for (var i = data.length - 1; i >= 1; i--) {
      var rowStatus = String(data[i][5] || "").trim();
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
 * SMS 대기 목록 자동 최신화
 */
function autoRefreshSmsLists() {
  try {
    checkLongTermAbsentees();         
    checkInactiveMembers();          
    checkInactivityDebuffAbsentees(); 
    
    Logger.log("SMS 발송 대기 목록 자동 최신화 완료");
  } catch (e) {
    Logger.log("SMS 자동 최신화 오류: " + e.toString());
  }
}

/**
 * 문자발송 상태 업데이트
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
 * 문자발송 내용 및 상태 동시 업데이트
 */
function updateSmsContentAndStatus(rowIdx, content, status) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("문자발송");
    if (!sheet) return { error: "시트를 찾을 수 없습니다." };
    
    sheet.getRange(rowIdx, 5).setValue(content); 
    sheet.getRange(rowIdx, 6).setValue(status);  
    return { success: true, message: "내용 수정 및 상태가 " + status + "(으)로 안전하게 업데이트되었습니다." };
  } catch (e) {
    return { error: "수정 및 업데이트 실패: " + e.toString() };
  }
}

/**
 * 7일 이상 미방문 회원 스캔 및 문자 자동 생성
 */
function checkLongTermAbsentees() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황");
    var logSheet = ss.getSheetByName("출석기록");
    var smsSheet = ss.getSheetByName("문자발송");
    
    if (!regSheet || !logSheet || !smsSheet) return { error: "필요한 시트(등록현황/출석기록/문자발송)가 없습니다." };
    
    var smsDataForCleanup = smsSheet.getDataRange().getValues();
    for (var i = smsDataForCleanup.length - 1; i >= 1; i--) {
      var category = String(smsDataForCleanup[i][3]);
      var status = String(smsDataForCleanup[i][5]).trim();
      if (category === "장기미방문" && status === "대기") {
        smsSheet.deleteRow(i + 1);
      }
    }
    
    var regData = regSheet.getDataRange().getDisplayValues();
    var logData = logSheet.getDataRange().getDisplayValues();
    var smsData = smsSheet.getDataRange().getDisplayValues();
    
    var regCols = getRegColumnIndices(regSheet);
    var logCols = getAttendanceColumnIndices(logSheet);
    
    var now = new Date();
    var sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(now.getDate() - 7);
    
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
    
    var hasPendingSmsMap = {};
    var lastSentSmsDateMap = {};
    for (var j = 1; j < smsData.length; j++) {
      var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
      var sCategory = smsData[j][3];
      var sStatus = smsData[j][5];
      var sDateStr = smsData[j][0]; 
      
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

    var memberMembershipNamesMap = {}; 
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

    var memberAttendanceTypeMap = {}; 
    for (var l = 1; l < logData.length; l++) {
      var lPhone = String(logData[l][logCols.phone] || "").replace(/[^0-9]/g, "");
      var lType = String(logData[l][logCols.type] || ""); 
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
    
    var count = 0;
    var addedNames = [];
    
    for (var k = 1; k < regData.length; k++) {
      var status = String(regData[k][regCols.status]).trim();
      if (status !== "진행중" && status !== "진행 중") continue;
      
      var phone = String(regData[k][regCols.phone] || "").replace(/[^0-9]/g, "");
      var name = regData[k][regCols.name];
      if (!phone) continue;
      
      var lastDate = lastAttendanceMap[phone];
      var cleanName = name.replace(/\d{4}$/, ""); 
      
      if (!lastDate || lastDate < sevenDaysAgo) {
        if (hasPendingSmsMap[phone]) continue;
        
        var lastSentSmsDate = lastSentSmsDateMap[phone];
        if (lastSentSmsDate) {
          if (lastDate && lastDate <= lastSentSmsDate) {
            var diffMs = now.getTime() - lastSentSmsDate.getTime();
            var daysSinceLastSms = diffMs / (1000 * 60 * 60 * 24);
            if (daysSinceLastSms < 7) {
              continue; 
            }
          }
        }
        
        var absentDays = 999; 
        if (lastDate) {
          absentDays = Math.floor((now.getTime() - lastDate.getTime()) / (1000 * 60 * 60 * 24));
        }
        
        var membership = regData[k][regCols.membership] || "회원권";
        var expire = regData[k][regCols.expire] || "미정";
        var remain = regData[k][regCols.remain] || "0";
        var expWarn = "\n🎫 회원님의 [" + membership + "] 만료일은 " + expire + "까지이며, 현재 잔여 횟수가 " + remain + "회 남아있습니다. 소중한 이용권이 마감되어 소멸되기 전에 꼭 오셔서 알차게 사용하시길 바랄게요! 😊";
        
        var formattedPhone = formatPhoneForSms(regData[k][regCols.phone]);
        
        var mTypes = memberMembershipNamesMap[phone] || [membership];
        var attStats = memberAttendanceTypeMap[phone] || { therapy: 0, jumping: 0 };
        var regClass = determineMemberClassInfo(mTypes, attStats);
        
        var msg = generateWellnessAiSms(cleanName, remain, absentDays, "장기미방문", formattedPhone, "sms", regClass, false, expire);
        
        var notiTitle = "회원님, 에너지를 다시 채워드릴게요! ❤️";
        if (absentDays >= 14 && absentDays <= 29) {
          notiTitle = "회원님의 활기찬 에너지가 무척 그립습니다! 😊";
        } else if (absentDays >= 30) {
          notiTitle = "회원님, 건강한 습관을 항상 응원합니다! 🌟";
        }
        
        var notiContent = generateWellnessAiSms(cleanName, remain, absentDays, "장기미방문", formattedPhone, "noti", regClass, false, expire) + expWarn;
        
        smsSheet.appendRow([
          Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd HH:mm"),
          cleanName,
          formattedPhone,
          "장기미방문",
          msg,
          "대기"
        ]);
        
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
    
    try {
      var questSheet = ss.getSheetByName("돌발퀘스트_목록");
      if (questSheet) {
        var questData = questSheet.getDataRange().getDisplayValues();
        var todayMidnight = new Date();
        todayMidnight.setHours(0,0,0,0);
        
        var localLogData = logData; 
        
        for (var q = 1; q < questData.length; q++) {
          var qType = questData[q][2];
          var qStatus = questData[q][7];
          
          if (qType === "글리코겐" && qStatus === "진행중") {
            var qPhone = String(questData[q][6]).replace(/[^0-9]/g, "");
            var qStartStr = questData[q][1];
            
            var qStartDate = new Date(qStartStr);
            qStartDate.setHours(0,0,0,0);
            var elapsedDays = Math.floor((todayMidnight.getTime() - qStartDate.getTime()) / (1000 * 60 * 60 * 24));
            
            var qStatusObj = getActiveQuestStatus(qPhone, ss, localLogData);
            if (qStatusObj && qStatusObj.glycogenQuest) {
              var qProgress = qStatusObj.glycogenQuest.progress;
              
              if (elapsedDays === 1 && qProgress === 0) {
                if (!hasSentNotificationToday(qPhone, "방어", "1일차 경고")) {
                  sendPersonalNotification(
                    qPhone,
                    "방어",
                    "🚨 글리코겐 방패 1일차 경고! 내일부터 2배 빡세게! 🔥",
                    "회원님, 어제 출석을 못 하셨네요! 😢 글리코겐이 몸의 지방으로 굳어버리기 전에 남은 이틀 동안 2배 더 빡세게 운동하셔야 합니다! 내일은 꼭 클럽에 나오셔서 활기차게 뛰고 글리코겐 방패를 채워보세요. 웰니스 코치가 기다리고 있습니다! 화이팅! 🏃‍♀️"
                  );
                }
              } else if (elapsedDays === 2 && qProgress <= 1) {
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
 * 14일 이상 미등록 회원 스캔 및 문자 자동 생성 (어드민 전용 수동 생성)
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
    
    var smsDataForCleanup = smsSheet.getDataRange().getValues();
    for (var i = smsDataForCleanup.length - 1; i >= 1; i--) {
      var category = String(smsDataForCleanup[i][3]);
      var status = String(smsDataForCleanup[i][5]).trim();
      if (category === "복귀권유" && status === "대기") {
        smsSheet.deleteRow(i + 1);
      }
    }
    
    var regData = regSheet.getDataRange().getDisplayValues();
    var smsData = smsSheet.getDataRange().getDisplayValues();
    var logData = logSheet ? logSheet.getDataRange().getDisplayValues() : [];
    
    var regCols = getRegColumnIndices(regSheet);
    var logCols = logSheet ? getAttendanceColumnIndices(logSheet) : null;
    
    var now = baseDateStr ? new Date(baseDateStr) : new Date();
    now.setHours(23, 59, 59, 999); 
    
    var memberAttendanceTypeMap = {}; 
    if (logSheet && logData.length >= 2 && logCols) {
      for (var l = 1; l < logData.length; l++) {
        var lPhone = String(logData[l][logCols.phone] || "").replace(/[^0-9]/g, "");
        var lType = String(logData[l][logCols.type] || ""); 
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
      
      if (status === "진행중" || status === "진행 중") {
        memberInfoMap[phone].isActive = true;
      }
      
      if (expireDateStr) {
        var expDate = new Date(expireDateStr);
        if (expDate > memberInfoMap[phone].lastExpire) {
          memberInfoMap[phone].lastExpire = expDate;
        }
      }
      
      memberInfoMap[phone].totalRemain += remain;
      
      if (membership) {
        memberInfoMap[phone].memberships.push(membership);
      }
    }
    
    var pendingSmsMap = {};
    var lastSentSmsDateMap = {};
    for (var j = 1; j < smsData.length; j++) {
      var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
      var sCategory = smsData[j][3];
      var sStatus = String(smsData[j][5]).trim();
      var sDateStr = smsData[j][0]; 
      
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
    
    var count = 0;
    var addedNames = [];
    
    var phones = Object.keys(memberInfoMap);
    for (var pIdx = 0; pIdx < phones.length; pIdx++) {
      var m = memberInfoMap[phones[pIdx]];
      
      if (!m.isActive && m.lastExpire.getTime() > 0 && m.lastExpire < now) {
        if (pendingSmsMap[phones[pIdx]]) continue;
        
        var lastSentSmsDate = lastSentSmsDateMap[phones[pIdx]];
        if (lastSentSmsDate) {
          var diffSmsMs = now.getTime() - lastSentSmsDate.getTime();
          var daysSinceLastSms = diffSmsMs / (1000 * 60 * 60 * 24);
          if (daysSinceLastSms < 7) {
            continue; 
          }
        }
        
        var diffMs = now.getTime() - m.lastExpire.getTime();
        var elapsedDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
        
        if (elapsedDays < 1) continue;
        
        var cleanName = m.name.replace(/\d{4}$/, ""); 
        var totalRemain = m.totalRemain;
        var formattedPhone = formatPhoneForSms(m.phoneRaw);
        
        var mTypes = m.memberships || [];
        var attStats = memberAttendanceTypeMap[phones[pIdx]] || { therapy: 0, jumping: 0 };
        var regClass = determineMemberClassInfo(mTypes, attStats);
        var hasNoAttendance = (!memberAttendanceTypeMap[phones[pIdx]] || (attStats.therapy === 0 && attStats.jumping === 0));
        
        var isAppActive = hasMemberActiveInApp(phones[pIdx], 3); 
        var targetChannel = "sms"; 
        
        if (elapsedDays >= 4 && elapsedDays <= 6) {
          if (isAppActive) {
            targetChannel = "noti"; 
          } else {
            targetChannel = "sms"; 
          }
        }
        
        var formattedExpireDate = (m.lastExpire && m.lastExpire.getTime() > 0) ? Utilities.formatDate(m.lastExpire, "GMT+9", "yyyy-MM-dd") : "기록없음 (미입력)";
        var msg = generateWellnessAiSms(cleanName, totalRemain, elapsedDays, "복귀권유", formattedPhone, targetChannel, regClass, hasNoAttendance, formattedExpireDate);
        
        if (targetChannel === "noti") {
          var notiTitle = "회원님, 남은 횟수(" + totalRemain + "회)를 지켜드릴게요! ❤️";
          if (totalRemain === 0) {
            notiTitle = "회원님, 함께 다져온 건강한 루틴을 계속 이어요! 😊";
          }
          sendPersonalNotification(formattedPhone, "복귀권유", notiTitle, msg);
        } else {
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
 * 정산용 진행중 회원 목록 가져오기
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
 * 개별 회원 출석 정산 문자 생성
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
 * 제미나이를 통한 개별 회원 맞춤형 한 달 출석 평가 코멘트 생성기
 */
function generateAttendanceReportAiComment(name, stats, prevSms, remain, expire, startDateStr, endDateStr) {
  try {
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
 * 4~6일 연속 결석으로 점수가 차감 중인 회원 추출 및 문자 생성
 */
function checkInactivityDebuffAbsentees() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    var logSheet = ss.getSheetByName("출석기록");
    var smsSheet = ss.getSheetByName("문자발송");
    
    if (!regSheet || !logSheet || !smsSheet) return { error: "필요한 시트(등록현황/출석기록/문자발송)가 없습니다." };
    
    var smsDataForCleanup = smsSheet.getDataRange().getValues();
    for (var i = smsDataForCleanup.length - 1; i >= 1; i--) {
      var category = String(smsDataForCleanup[i][3]);
      var status = String(smsDataForCleanup[i][5]).trim();
      if (category === "출석디버프" && status === "대기") {
        smsSheet.deleteRow(i + 1);
      }
    }
    
    var regData = regSheet.getDataRange().getDisplayValues();
    var logData = logSheet.getDataRange().getValues();
    var smsData = smsSheet.getDataRange().getDisplayValues();
    
    var regCols = getRegColumnIndices(regSheet);
    var logCols = getAttendanceColumnIndices(logSheet);
    
    var now = new Date();
    
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
    
    var hasPendingSmsMap = {};
    var lastSentSmsDateMap = {};
    for (var j = 1; j < smsData.length; j++) {
      var sPhone = String(smsData[j][2] || "").replace(/[^0-9]/g, "");
      var sCategory = smsData[j][3];
      var sStatus = smsData[j][5];
      var sDateStr = smsData[j][0]; 
      
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
      if (!lastDate) continue; 
      
      var midnightLast = new Date(lastDate.getTime());
      midnightLast.setHours(0,0,0,0);
      
      var inactiveDays = calculateInactiveDays(midnightLast, midnightNow);
      
      if (inactiveDays >= 4 && inactiveDays <= 6) {
        if (hasPendingSmsMap[phone]) continue;
        
        var lastSentSmsDate = lastSentSmsDateMap[phone];
        if (lastSentSmsDate) {
          if (lastDate && lastDate <= lastSentSmsDate) {
            var diffMs = now.getTime() - lastSentSmsDate.getTime();
            var daysSinceLastSms = diffMs / (1000 * 60 * 60 * 24);
            if (daysSinceLastSms < 7) {
              continue; 
            }
          }
        }
        
        var cleanName = name.replace(/\d{4}$/, ""); 
        var penaltyVal = (inactiveDays - 3) * 100;
        
        var notiTitle = "🚨 웰니스 점수 방전 경고! 홈 케어 미션 🔥";
        var notiContent = cleanName + " 회원님! 운동을 쉬신 지 벌써 " + inactiveDays + "일이 지나, 아쉽게도 웰니스 누적 점수가 매일 100 EXP씩 방전(감점)되기 시작했어요! 😭 누적 -" + penaltyVal + " EXP 상태입니다.\n\n하지만 걱정 마세요! 오늘 클럽에 오셔서 신나게 점핑 뛰고 출석체크만 쾅! 하시면 깎였던 모든 점수와 순위가 즉시 100% 마법처럼 전부 복구(부활)됩니다! ✨\n\n만약 방문이 어려우시다면 홈 케어로 가벼운 스트레칭이나 물 2L 마시기 미션이라도 실천해보시는 건 어떨까요? 오늘 꼭 앱에서 일상 기록을 남겨 에너지를 지키세요! ❤️";
        
        var isActive = isAppActiveDuringAbsence(phone, lastDate, now);
        
        try {
          if (!hasSentNotificationToday(phone, "디버프", "웰니스 점수 방전 경고")) {
            sendPersonalNotification(phone, "디버프", notiTitle, notiContent);
          }
        } catch (notiErr) {
          Logger.log("디버프 인앱 쪽지 생성 실패 (" + cleanName + "): " + notiErr.toString());
        }
        
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

/**
 * 9. 기타 어드민 도구 & 지니월드 관리자 설정 API
 */

// [이동 완료] fixAllPhoneNumbersInSheet 함수는 Common_Utils.gs로 공통화 이관되었습니다.

// [이동 완료] checkMemberTicket 함수는 Common_Utils.gs로 공통화 이관되었습니다.

