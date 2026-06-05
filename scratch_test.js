const fs = require('fs');
const path = require('path');

const root = 'g:/내 드라이브/nohyung-jumping-web';

// 다른 파일로 이관된 함수들을 파악하기 위해 각 파일에서 함수명을 추출합니다.
function getDefinedFunctions(filePath) {
  if (!fs.existsSync(filePath)) return new Set();
  const content = fs.readFileSync(filePath, 'utf8');
  const funcRegex = /function\s+(\w+)\s*\(/g;
  const funcs = new Set();
  let match;
  while ((match = funcRegex.exec(content)) !== null) {
    funcs.add(match[1]);
  }
  return funcs;
}

const adminFuncs = getDefinedFunctions(path.join(root, 'Admin.gs'));
const bridgeFuncs = getDefinedFunctions(path.join(root, 'Bridge_Login.gs'));
const sanctumFuncs = getDefinedFunctions(path.join(root, 'Sanctum.gs'));

// 이관 완료된 전체 함수 집합
const migratedFuncs = new Set([...adminFuncs, ...bridgeFuncs, ...sanctumFuncs]);

// 보존해야 할 코어/공용 함수들 (이것들은 Code.gs에 반드시 유지해야 함)
const coreKeepFuncs = new Set([
  'onEdit',
  'syncClubRecord',
  'setupDatabase',
  'forceUpdateAllHeaders',
  'autoExpireMemberships',
  'migrateRegistrationSheet',
  'doGet',
  'doPost',
  'handleRequest',
  'include',
  'getScriptUrl',
  'getMemberSheetColumnIndices',
  'getRegColumnIndices',
  'getAttendanceColumnIndices',
  'formatTimeSafely',
  'parseDateTimeSafely',
  'normalizePhoneDigits',
  'formatPhoneNumber',
  'parseWeightSafely',
  'normalizeDateStr',
  'loadPublicHolidaysOnce',
  'loadFlashHolidayMapOnce',
  'isKoreanPublicHoliday',
  'isCenterHoliday',
  'calculateInactiveDays',
  'submitSalesRecord',
  'generateSmsContent',
  'getFlashSettingByDate',
  'saveFlashSetting',
  'getAllFlashSettings',
  'deleteFlashSettingByRow',
  'fixAllPhoneNumbersInSheet',
  'checkMemberTicket',
  'getCompiledMemberRegistry',
  'getMemberRenewalData',
  'SeededRandom',
  'calculateInbodyScoreHelper',
  'getThuStartOfWeekLocal',
  'getWeekStringLocal',
  'getFriendlyName',
  'get4thWednesday',
  'isDateInLastWeekMonToWed',
  'determineMemberClassInfo',
  'getMemberIDListWithTicket',
  'getActiveUserPhones',
  'getOrCreateFolder',
  'getMemberIDList'
]);

// 1. Code.gs 로드
const codeGsPath = path.join(root, 'Code.gs');
let codeContent = fs.readFileSync(codeGsPath, 'utf8');

// 정밀하게 함수 단위를 파싱하기 위한 초간단 파서
// 중괄호 {} 쌍 매칭을 이용해 각 함수의 시작 위치와 끝 위치를 구합니다.
function stripMigratedFunctions(code) {
  let output = '';
  let pos = 0;
  
  // function 함수명 ( 형식의 정규식
  const funcRegex = /function\s+(\w+)\s*\(/g;
  let match;
  
  const funcPositions = [];
  while ((match = funcRegex.exec(code)) !== null) {
    const funcName = match[1];
    const startIdx = match.index;
    funcPositions.push({ name: funcName, start: startIdx });
  }
  
  // 역순으로 지워나가야 인덱스가 꼬이지 않습니다.
  for (let i = funcPositions.length - 1; i >= 0; i--) {
    const item = funcPositions[i];
    
    // 만약 이관된 함수이고, 코어 보존 함수가 아니라면 삭제 진행
    if (migratedFuncs.has(item.name) && !coreKeepFuncs.has(item.name)) {
      // 함수의 시작 지점부터 닫는 중괄호 } 까지의 끝 지점을 찾습니다.
      let braceCount = 0;
      let started = false;
      let endIdx = -1;
      
      for (let j = item.start; j < code.length; j++) {
        const char = code[j];
        if (char === '{') {
          braceCount++;
          started = true;
        } else if (char === '}') {
          braceCount--;
          if (started && braceCount === 0) {
            endIdx = j + 1;
            break;
          }
        }
      }
      
      if (endIdx !== -1) {
        // [삭제 및 주석 대체]
        const before = code.substring(0, item.start);
        const after = code.substring(endIdx);
        code = before + `// [이관 완료] function ${item.name} 이(가) 다른 모듈로 이관되어 본 파일에서 안전하게 제거되었습니다.\n` + after;
        console.log(`🧹 Stripped migrated function: ${item.name}`);
      }
    }
  }
  return code;
}

const cleanedCode = stripMigratedFunctions(codeContent);
fs.writeFileSync(codeGsPath, cleanedCode, 'utf8');
console.log('Code.gs cleanup completed! Remaining lines:', cleanedCode.split('\n').length);
