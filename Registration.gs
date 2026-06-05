/**
 * ==========================================
 * 노형점핑 & 체온테라피 통합 관리 웹앱
 * [Registration.gs] 회원 가입 및 등록 전용 API 엔진
 * ==========================================
 */

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

/**
 * 신규 회원 가입 등록 API
 */
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
    
    var curRemainVal = Number(existingRemain) || 0;
    
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
