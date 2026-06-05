/**
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * Kiosk Attendance Module
 */

// [이동 완료] getCompiledMemberRegistry 함수는 Common_Utils.gs로 공통화 이관되었습니다.

/**
 * 백엔드 API: 회원 검색 (태블릿 출석체크용)
 */
function searchMemberByPin(pinStr) {
  try {
    var pin = String(pinStr || "").trim();
    if (pin.length < 4) return { error: "뒷자리 4자리를 정확히 입력해주세요." };
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var registry = getCompiledMemberRegistry(ss);
    
    var matched = [];
    for (var i = 0; i < registry.length; i++) {
      if (registry[i].phoneClean.slice(-4) === pin) {
        matched.push(registry[i]);
      }
    }
    
    if (matched.length === 0) {
      // [BUGFIX] 캐시 강제 무효화 후 재조회
      var cache = CacheService.getScriptCache();
      cache.remove("v45_member_registry");
      registry = getCompiledMemberRegistry(ss);
      
      for (var i = 0; i < registry.length; i++) {
        if (registry[i].phoneClean.slice(-4) === pin) {
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

/**
 * 백엔드 API: 출석 처리 및 차감 (태블릿 출석체크용)
 */
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
      }
      selectedPasses.push({ pass: pass, amount: deductAmount });
      
    } else if (type === '테라피') {
      var pass = null;
      for (var pIdx = 0; pIdx < activePasses.length; pIdx++) {
        var p = activePasses[pIdx];
        if (p.membershipType.indexOf("테라피") !== -1 || p.membershipType.indexOf("패키지") !== -1) {
          pass = p;
          break;
        }
      }
      if (!pass) pass = activePasses[0];
      if (!pass) return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };
      
      var rule = configRules[pass.membershipType];
      var deductAmount = 1;
      
      if (rule && rule.Therapy !== "불가") {
        deductAmount = Number(rule.Therapy);
      } else if (!rule) {
        if (pass.membershipType.indexOf("테라피") !== -1 || pass.membershipType.indexOf("패키지") !== -1) {
          deductAmount = 1;
        } else {
          return { error: "이용 가능한 회원권이 없습니다. 다시 확인해주세요..." };
        }
      }
      selectedPasses.push({ pass: pass, amount: deductAmount });
      
    } else if (type === '복합') {
      // [복합 모드] 한 회원권에서 일괄 차감 또는 복수 회원권 교차 차감 지원
      var pass = null;
      for (var pIdx = 0; pIdx < activePasses.length; pIdx++) {
        var p = activePasses[pIdx];
        if (p.membershipType.indexOf("복합") !== -1 || p.membershipType.indexOf("패키지") !== -1) {
          pass = p;
          break;
        }
      }
      
      if (pass) {
        // 복합 단일권이 존재하는 경우
        var rule = configRules[pass.membershipType];
        var deductAmount = 1;
        if (rule && rule.Combo !== "불가") {
          deductAmount = Number(rule.Combo);
        }
        selectedPasses.push({ pass: pass, amount: deductAmount });
      } else {
        // 복합 단일권이 없어 각각 1장씩 교차 차감하는 경우
        var jumpPass = null;
        var therapyPass = null;
        
        for (var pIdx = 0; pIdx < activePasses.length; pIdx++) {
          var p = activePasses[pIdx];
          if (!jumpPass && (p.membershipType.indexOf("점핑") !== -1 || p.membershipType.indexOf("월권") !== -1 || p.membershipType.indexOf("운동만") !== -1)) {
            jumpPass = p;
          }
          if (!therapyPass && (p.membershipType.indexOf("테라피") !== -1)) {
            therapyPass = p;
          }
        }
        
        if (jumpPass && therapyPass) {
          var jRule = configRules[jumpPass.membershipType];
          var tRule = configRules[therapyPass.membershipType];
          
          var jDeduct = (jRule && jRule.Jumping !== "불가") ? Number(jRule.Jumping) : 1;
          var tDeduct = (tRule && tRule.Therapy !== "불n") ? Number(tRule.Therapy) : 1;
          
          selectedPasses.push({ pass: jumpPass, amount: jDeduct });
          selectedPasses.push({ pass: therapyPass, amount: tDeduct });
        } else {
          // 최후의 보루: 그냥 가지고 있는 첫 번째 회원권에서 차감
          selectedPasses.push({ pass: activePasses[0], amount: 1 });
        }
      }
    } else if (isBonus) {
      if (bonusCount <= 0) return { error: "사용 가능한 보너스 횟수가 없습니다." };
    }

    // 4. 차감 실행 및 로그 기록
    var logDateStr = Utilities.formatDate(now, "GMT+9", "yyyy-MM-dd");
    var logTimeStr = Utilities.formatDate(now, "GMT+9", "HH:mm:ss");
    var firstMemberName = activePasses.length > 0 ? activePasses[0].name : (matchedMember ? matchedMember.name : "회원");
    var usedPassName = "";
    var usedExpireDate = "-";
    
    // [트랜잭션 락] 안전한 차감 처리를 위해 락 획득
    var lock = LockService.getScriptLock();
    lock.waitLock(5000); // 최대 5초 대기
    
    var nextCount = "";
    if (isBonus) {
      // 보너스 차감
      if (mRowIdx !== -1) {
        var newBonus = Math.max(0, bonusCount - 1);
        memberSheet.getRange(mRowIdx, 10).setValue(newBonus); // 10번째 열: 보너스횟수 (J열)
        nextCount = String(newBonus);
        usedPassName = "보너스권";
        
        logSheet.appendRow([
          logDateStr,
          logTimeStr,
          firstMemberName,
          "'" + phoneStr,
          "보너스",
          bonusCount,
          -1,
          newBonus,
          "보너스 1회 차감",
          "", // 참여클래스 비어둠
          "0",
          "출석",
          "", // 퇴실시간 비어둠
          ""  // 비고 비어둠
        ]);
      }
    } else {
      // 일반/복합 회원권 차감
      for (var sIdx = 0; sIdx < selectedPasses.length; sIdx++) {
        var item = selectedPasses[sIdx];
        var pass = item.pass;
        var amount = item.amount;
        
        var currentVal = regSheet.getRange(pass.rowIdx, cols.remain + 1).getValue(); 
        var isUnlimited = (currentVal === "(무제한)" || currentVal === "무제한" || String(currentVal).trim() === "");
        var remain = isUnlimited ? 9999 : Number(currentVal);
        var nextRemain = isUnlimited ? "무제한" : Math.max(0, remain - amount);
        
        if (!isUnlimited) {
          regSheet.getRange(pass.rowIdx, cols.remain + 1).setValue(nextRemain);
        }
        
        nextCount = isUnlimited ? "무제한" : String(nextRemain);
        usedPassName = pass.membershipType;
        usedExpireDate = pass.expireDate || "-";
        
        // 만약 차감 후 잔여 횟수가 0회라면 상태를 자동으로 "마감" 처리합니다. (기한 무제한 제외)
        if (!isUnlimited && nextRemain <= 0) {
          regSheet.getRange(pass.rowIdx, cols.status + 1).setValue("마감");
        }
        
        // 출석기록 작성
        logSheet.appendRow([
          logDateStr,
          logTimeStr,
          firstMemberName,
          "'" + phoneStr,
          usedPassName,
          currentVal, // 차감 전 잔여
          isUnlimited ? 0 : -amount, // 변동량
          nextRemain, // 차감 후 잔여
          type + " 출석 차감" + (pass.memo ? " (" + pass.memo + ")" : ""),
          "", // 참여클래스 비어둠
          "0",
          "출석",
          "", // 퇴실시간 비어둠
          ""  // 비고 비어둠
        ]);
      }
    }
    
    lock.releaseLock();
    
    // [캐시 강제 갱신] 차감 완료 후 10분 캐시를 즉각 삭제하여 다음 조회 시 리얼타임 데이터가 보장되도록 결계 구축
    try {
      var scriptCache = CacheService.getScriptCache();
      scriptCache.remove("v45_member_registry");
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

// [이동 완료] processKioskCheckout 함수는 Admin.gs로 이관되었습니다.

