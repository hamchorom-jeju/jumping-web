function checkMemberTicket(name, phoneRaw) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var regSheet = ss.getSheetByName("등록 현황");
    var memberSheet = ss.getSheetByName("회원명단");
    
    var phoneOnly = phoneRaw.replace(/[^0-9]/g, "");
    var today = new Date();
    today.setHours(0,0,0,0);
    
    var hasValidTicket = false;
    var tickets = [];

    // 1. 등록 현황 체크 (테라피권, 점핑30회, 점핑50회 등)
    if (regSheet) {
      var regData = regSheet.getDataRange().getValues();
      var cols = getRegColumnIndices(regSheet);
      for (var i = cols.headerRow + 1; i < regData.length; i++) {
        var rName = String(regData[i][cols.name]).trim();
        var rPhone = String(regData[i][cols.phone]).replace(/[^0-9]/g, "");
        
        if (rName === name && rPhone === phoneOnly) {
          var type = String(regData[i][cols.membership]).trim();
          var remain = Number(regData[i][cols.remain]);
          var expireRaw = regData[i][cols.expire];
          var expireDate = expireRaw instanceof Date ? expireRaw : new Date(expireRaw);
          var status = String(regData[i][cols.status]).trim();

          // 테라피/점핑 회원권 여부 및 유효성 확인
          if (type.indexOf("테라피") !== -1 || type.indexOf("점핑30") !== -1 || type.indexOf("점핑50") !== -1) {
             if (status !== "마감" && status !== "정지" && remain > 0 && expireDate >= today) {
               hasValidTicket = true;
               tickets.push(type + "(" + remain + "회)");
             }
          }
        }
      }
    }

    // 2. 보너스권 체크 (회원명단 J열 - Index 9)
    if (memberSheet) {
      var mData = memberSheet.getDataRange().getValues();
      for (var j = 1; j < mData.length; j++) {
        var mName = String(mData[j][1]).trim();
        var mPhone = String(mData[j][2]).replace(/[^0-9]/g, "");
        if (mName === name && mPhone === phoneOnly) {
          var bonus = Number(mData[j][9]); 
          if (bonus > 0) {
            hasValidTicket = true;
            tickets.push("보너스권(" + bonus + "회)");
          }
        }
      }
    }

    return {
      success: true,
      hasValidTicket: hasValidTicket,
      tickets: tickets.length > 0 ? tickets.join(", ") : "사용 가능한 회원권 없음"
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}
