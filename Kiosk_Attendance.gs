/**
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * Kiosk Attendance Module
 */

// ──────────────────────────────────────────────
// 3. 백엔드 API: 회원 검색 (태블릿 출석체크용)
// ──────────────────────────────────────────────
/**
 * [perf] 활성 회원 목록을 취합하여 구글 Apps Script 캐시 서비스에 10분 동안 저장하는 최적화 헬퍼
 */
function getCompiledMemberRegistry(ss) {
  var cache = CacheService.getScriptCache();
  var cacheKey = "v45_member_registry";
  var cached = cache.get(cacheKey);
  if (cached) {
    try {
      var parsed = JSON.parse(cached);
      if (parsed && parsed.length > 0) {
        return parsed;
      }
    } catch(e) {}
  }
  
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
  if (!regSheet) return [];
  
  var data = regSheet.getDataRange().getDisplayValues();
  var cols = getRegColumnIndices(regSheet);
  
  var memberSheet = ss.getSheetByName("회원명단");
  var mData = memberSheet ? memberSheet.getDataRange().getDisplayValues() : [];
  
  var memberMap = {};
  
  for (var i = 1; i < data.length; i++) {
    var phoneRaw = data[i][cols.phone];
    var phoneClean = formatPhoneNumber(phoneRaw).replace(/[^0-9]/g, ""); 
    var status = String(data[i][cols.status] || "").trim(); 
    if (status === "진행중" || status === "진행 중" || status === "마감") {
      if (!memberMap[phoneClean]) {
        var bonus = "0";
        var mRowIdx = -1;
        if (mData.length > 0) {
          for (var mIdx = 1; mIdx < mData.length; mIdx++) {
            var mPhone = formatPhoneNumber(mData[mIdx][2]).replace(/[^0-9]/g, "");
            if (mPhone === phoneClean) {
              bonus = String(mData[mIdx][9] || "0"); // 9번 컬럼: 보너스횟수
              mRowIdx = mIdx + 1;
              break;
            }
          }
        }
        
        memberMap[phoneClean] = {
          name: String(data[i][cols.name] || "이름없음"),
          phone: phoneRaw,
          bonusCount: bonus,
          mRowIdx: mRowIdx,
          passes: []
        };
      }
      
      memberMap[phoneClean].passes.push({
        membershipType: String(data[i][cols.membership] || "일반"),
        expireDate: data[i][cols.expire],
        remainCount: String(data[i][cols.remain] || "0"),
        rowIdx: i + 1,
        memo: data[i][10], // K열 비고(이월딱지)
        status: status
      });
    }
  }
  
  var registryList = [];
  var keys = Object.keys(memberMap);
  for (var j = 0; j < keys.length; j++) {
    var m = memberMap[keys[j]];
    var allExpired = m.passes.every(function(p) { return p.status === "마감"; });
    
    registryList.push({
      name: m.name,
      phone: m.phone,
      membershipType: m.passes.map(function(p) { return p.membershipType; }).join(" / "),
      expireDate: m.passes[0].expireDate,
      remainCount: m.passes[0].remainCount,
      bonusCount: m.bonusCount,
      mRowIdx: m.mRowIdx,
      allPasses: m.passes,
      phoneClean: keys[j],
      isExpired: allExpired
    });
  }
  
  try {
    cache.put(cacheKey, JSON.stringify(registryList), 600); // 10분 캐시
  } catch(e) {}
  
  return registryList;
}

// ──────────────────────────────────────────────
// 3. 백엔드 API: 회원 검색 (태블릿 출석체크용)
// ──────────────────────────────────────────────
function searchMemberByPin(pinStr) {
  try {
    var pin = String(pinStr || "").trim();
    if (pin.length < 4) return { error: "뒷자리 4자리를 정확히 입력해주세요." };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var registry = getCompiledMemberRegistry(ss);
    
    var matched = [];
    for (var i = 0; i < registry.length; i++) {
      if (registry[i].phoneClean.slice(-4) === pin && !registry[i].isExpired) {
        matched.push(registry[i]);
      }
    }
    
    if (matched.length === 0) {
      // [BUGFIX] 혹시 원장님이 수동으로 시트를 고친 경우, 10분 캐시 때문에 안 나왔을 가능성 대비 캐시 강제 무효화 후 스프레드시트 재로드!
      var cache = CacheService.getScriptCache();
      cache.remove("v45_member_registry");
      registry = getCompiledMemberRegistry(ss);
      
      for (var i = 0; i < registry.length; i++) {
        if (registry[i].phoneClean.slice(-4) === pin && !registry[i].isExpired) {
          matched.push(registry[i]);
        }
      }
      
      if (matched.length === 0) {
        return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };
      }
    }
    
    // 추가: 테라피 입실 중인지 예약DB 확인 (퇴실 버튼 표시용)
    var isTherapyActive = false;
    try {
      var resSheet = ss.getSheetByName("예약DB");
      if (resSheet) {
        var lastRow = resSheet.getLastRow();
        if (lastRow > 1) {
          // [perf] 오늘 예약 건만 확인하기 위해 최근 100행만 정밀 탐색하여 속도 향상 (100배 빠름)
          var checkStartRow = Math.max(2, lastRow - 100);
          var checkRowCount = lastRow - checkStartRow + 1;
          var resData = resSheet.getRange(checkStartRow, 1, checkRowCount, 11).getDisplayValues();
          
          var todayFormatted = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
          var todayNum = todayFormatted.replace(/[^0-9]/g, "");
          
          for (var r = 0; r < resData.length; r++) {
            var rDateStr = String(resData[r][3]).replace(/[^0-9]/g, "");
            var rPhone = String(resData[r][1]).replace(/[^0-9]/g, "");
            var rStatus = String(resData[r][9]);
            
            if (rDateStr === todayNum && rPhone.slice(-4) === pin && rStatus === "테라피 진행중") {
              isTherapyActive = true;
              break;
            }
          }
        }
      }
    } catch(e) {}
    
    return { success: true, members: matched, isTherapyActive: isTherapyActive };
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
    
    // [perf] 캐시된 회원 정보 및 설정 캐시 로드 (0.02초)
    var registry = getCompiledMemberRegistry(ss);
    var cleanPhone = phoneStr.replace(/[^0-9]/g, "");
    var matchedMember = null;
    for (var m = 0; m < registry.length; m++) {
      if (registry[m].phoneClean === cleanPhone) {
        matchedMember = registry[m];
        break;
      }
    }
    
    var activePasses = [];
    var bonusCount = 0;
    var mRowIdx = -1;
    
    if (matchedMember) {
      bonusCount = Number(matchedMember.bonusCount) || 0;
      mRowIdx = matchedMember.mRowIdx;
      for (var p = 0; p < matchedMember.allPasses.length; p++) {
        var pass = matchedMember.allPasses[p];
        if (pass.status === "진행중" || pass.status === "진행 중") {
          activePasses.push({
            rowIdx: pass.rowIdx,
            name: matchedMember.name,
            membershipType: pass.membershipType,
            remainCount: pass.remainCount,
            memo: pass.memo
          });
        }
      }
    }
    
    var regSheet = ss.getSheetByName("등록 현황") || ss.getSheetByName("등록현황");
    
    if (activePasses.length === 0 && bonusCount <= 0) {
      return { 
        error: "이용 가능한 회원권이 없습니다. (입력번호: " + phoneStr + ", cleanPhone: " + cleanPhone + ", 매칭회원 찾았는지: " + (matchedMember ? "찾음, 회원권개수: " + matchedMember.allPasses.length + ", 보너스: " + matchedMember.bonusCount : "못찾음 (전체명단 수: " + registry.length + "명)") + ")"
      };
    }

    // [perf] 설정 규칙 24시간 캐싱 처리
    var configRules = null;
    var cache = CacheService.getScriptCache();
    var cachedConfig = cache.get("v44_config_rules");
    if (cachedConfig) {
      try { configRules = JSON.parse(cachedConfig); } catch(e) {}
    }
    if (!configRules) {
      var configSheet = ss.getSheetByName("설정");
      var cData = configSheet ? configSheet.getDataRange().getValues() : [];
      configRules = {};
      for (var j = 1; j < cData.length; j++) {
        configRules[cData[j][0]] = { Jumping: cData[j][3], Therapy: cData[j][4], Combo: cData[j][5] };
      }
      try { cache.put("v44_config_rules", JSON.stringify(configRules), 86400); } catch(e) {}
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

        if (type === '복합') {
          comboLogInfo.prev = (comboLogInfo.prev === 0) ? prevCount : comboLogInfo.prev;
          comboLogInfo.change += deductAmount;
          comboLogInfo.remain = nextCount;
          comboLogInfo.reason += (comboLogInfo.reason ? " + " : "") + item.reason;
        } else {
          logSheet.appendRow([dateStr, timeStr, firstMemberName, phoneStr, pass.membershipType, prevCount, "-" + deductAmount, nextCount, item.reason, "", "", "입실", "", ""]);
        }
      }
    }

    if (type === '복합' && selectedPasses.length > 0) {
      logSheet.appendRow([dateStr, timeStr, firstMemberName, phoneStr, "복합", comboLogInfo.prev, "-" + comboLogInfo.change, comboLogInfo.remain, comboLogInfo.reason, "", "", "입실", "", ""]);
      nextCount = comboLogInfo.remain; 
    }

    // --- [예약DB 실시간 연동] 테라피/복합 출석 시 예약 상태 업데이트 ---
    if (type === '테라피' || type === '복합' || type === '보너스') {
      try {
        var resSheet = ss.getSheetByName("예약DB");
        if (resSheet) {
          var lastRow = resSheet.getLastRow();
          if (lastRow > 1) {
            var checkStartRow = Math.max(2, lastRow - 100);
            var checkRowCount = lastRow - checkStartRow + 1;
            var resData = resSheet.getRange(checkStartRow, 1, checkRowCount, 11).getDisplayValues();
            
            var todayFormatted = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
            var todayParts = todayFormatted.match(/\d+/g);
            var todayNum = todayParts[0] + (todayParts[1].length === 1 ? "0" : "") + todayParts[1] + (todayParts[2].length === 1 ? "0" : "") + todayParts[2];
            var phoneClean = String(phoneStr || "").replace(/[^0-9]/g, "");
            
            for (var r = 0; r < resData.length; r++) {
              var rDateRaw = resData[r][3];
              var rDateParts = String(rDateRaw).match(/\d+/g);
              if (!rDateParts || rDateParts.length < 3) continue;
              var rDateNum = rDateParts[0] + (rDateParts[1].length === 1 ? "0" : "") + rDateParts[1] + (rDateParts[2].length === 1 ? "0" : "") + rDateParts[2];
              
              var rPhone = String(resData[r][1]).replace(/[^0-9]/g, "");
              var isPhoneMatch = (rPhone === phoneClean) || 
                                 (rPhone.length >= 8 && phoneClean.length >= 8 && rPhone.slice(-8) === phoneClean.slice(-8));

              if (rDateNum === todayNum && isPhoneMatch) {
                var status = String(resData[r][9]);
                if (status.indexOf("예약") !== -1 || status.indexOf("테라피") !== -1) {
                  var actualRowIdx = checkStartRow + r;
                  resSheet.getRange(actualRowIdx, 10).setValue("테라피중");
                  resSheet.getRange(actualRowIdx, 11).setValue(Utilities.formatDate(now, "GMT+9", "HH:mm:ss"));
                  break;
                }
              }
            }
          }
        }
      } catch (resErr) {
        console.error("예약DB 연동 오류: " + resErr.toString());
      }
    }

    var usedPassName = "";
    var usedExpireDate = "";
    if (type === '보너스') {
      usedPassName = "보너스권";
      usedExpireDate = "-";
    } else if (type === '복합') {
      usedPassName = "복합(점핑30회)";
      if (selectedPasses.length > 0 && selectedPasses[0].pass) {
        var pIdx = selectedPasses[0].pass.rowIdx;
        usedExpireDate = regSheet.getRange(pIdx, 7).getDisplayValue(); 
      }
    } else if (selectedPasses.length > 0 && selectedPasses[0].pass) {
      usedPassName = selectedPasses[0].pass.membershipType;
      var pIdx = selectedPasses[0].pass.rowIdx;
      usedExpireDate = regSheet.getRange(pIdx, 7).getDisplayValue();
    }

    clearUserDashboardCache(phoneStr);
    try {
      var cache = CacheService.getScriptCache();
      cache.remove("v45_member_registry");
    } catch(e) {}

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
