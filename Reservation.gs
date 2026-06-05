/**
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * Therapy Room Reservation Portal Module
 */

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
    var status = String(row[9] || "").replace(/\s/g, ""); 
    
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
    var seen = {};

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
