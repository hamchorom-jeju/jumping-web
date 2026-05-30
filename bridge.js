/**
 * Nohyung Jumping ERP - Vercel Bridge Script
 * (C) 2024 Antigravity AI
 */

const API_URL = "/api/gas";

if (typeof google === 'undefined' || !google.script) {
  console.log("🚀 Vercel Bridge Active: Mocking google.script.run");
  
  const bridgeTarget = {
    successHandler: null,
    failureHandler: null,
    withSuccessHandler: function(callback) {
      this.successHandler = callback;
      return this;
    },
    withFailureHandler: function(callback) {
      this.failureHandler = callback;
      return this;
    }
  };

  window.google = {
    script: {
      run: new Proxy(bridgeTarget, {
        get: function(target, prop) {
          if (prop in target) {
            return target[prop];
          }

          return function(...args) {
            const sHandler = target.successHandler;
            const fHandler = target.failureHandler;
            
            // 핸들러 초기화 (다음 호출을 위해)
            target.successHandler = null;
            target.failureHandler = null;

            // [v45.10] 만약 로컬 환경(file://, localhost, 127.0.0.1)이라면, 실서버 통신 대신 가상 모의(Mock) 데이터 즉시 반환!
            const isLocalEnv = window.location.protocol === 'file:' || window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1';
            if (isLocalEnv) {
              console.log(`🎮 [Local Mock Mode] Intercepted: ${prop}`, args);
              setTimeout(() => {
                let mockResult = { success: true };
                
                if (prop === 'searchMembersByDigits') {
                  mockResult = {
                    success: true,
                    members: [
                      { name: "홍길동", phone: "010-1234-5678" },
                      { name: "이장님", phone: "010-9999-9999" }
                    ]
                  };
                } else if (prop === 'getUserDashboardData') {
                  mockResult = {
                    success: true,
                    name: "이장님",
                    tier: "🏆 마스터 모험가",
                    exp: 750,
                    level: 5,
                    totalScore: 750,
                    rank: 1,
                    doneList: ["아침 식단"],
                    weeklyTargets: { health: 1500, perf: 1000, def: 500 },
                    monthlyTargets: { health: 6000, perf: 4000, def: 2000 },
                    villageSettings: {
                      weather: "blossom",
                      bgmEnabled: "true",
                      bgmUrl: "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3"
                    },
                    pillarNotice: [
                      { title: "노형 빌리지 리뉴얼 완공! 🏰", content: "따사로운 봄바람과 함께 벚꽃 마법이 시작되었습니다. 🌸 우측 하단의 BGM 플레이어를 켜고 낭만 넘치는 음악을 감상해 보세요!" },
                      { title: "33 챌린지 시즌 4주차 결산 안내 🎁", content: "모든 모험가 여러분의 노고를 기리며, 오늘 밤 특별 보상이 지급될 예정입니다!" },
                      { title: "20시 이후 금식 엄수! 🍳", content: "완벽한 성취를 위해 마지막 식단 원칙을 지켜주세요." }
                    ]
                  };
                } else if (prop === 'getArchiveFeed') {
                  let stored = sessionStorage.getItem("mock_feed");
                  if (!stored) {
                    const defaultMockFeed = [
                      { name: "이장님", item: "아침 식단", comment: "아침은 역시 사과와 미온수!", type: "식단", photoId: "", date: "2026-05-18", time: "08:30" },
                      { name: "홍길동", item: "오운완", comment: "오늘도 땀 흘리며 점핑 완주!", type: "퀘스트", photoId: "", date: "2026-05-18", time: "19:45" }
                    ];
                    sessionStorage.setItem("mock_feed", JSON.stringify(defaultMockFeed));
                    stored = JSON.stringify(defaultMockFeed);
                  }
                  mockResult = {
                    success: true,
                    items: JSON.parse(stored),
                    totalPages: 1,
                    currentPage: 1
                  };
                } else if (prop === 'getVillageSettings') {
                  const storedWeather = localStorage.getItem("mock_weather") || "blossom";
                  const storedBgmEnabled = localStorage.getItem("mock_bgm_enabled") || "true";
                  const storedBgmUrl = localStorage.getItem("mock_bgm_url") || "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3";
                  
                  const bgm_sun = localStorage.getItem("mock_bgm_sun") || "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3";
                  const bgm_rain = localStorage.getItem("mock_bgm_rain") || "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-8.mp3";
                  const bgm_snow = localStorage.getItem("mock_bgm_snow") || "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-10.mp3";
                  const bgm_blossom = localStorage.getItem("mock_bgm_blossom") || "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-3.mp3";
                  const bgm_leaves = localStorage.getItem("mock_bgm_leaves") || "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-5.mp3";
                  
                  let weather = storedWeather;
                  let bgmUrl = storedBgmUrl;
                  
                  if (weather === "auto") {
                    weather = "blossom"; // 로컬 프리뷰에서는 제주의 봄철 벚꽃 날씨와 BGM으로 시뮬레이션!
                    bgmUrl = bgm_blossom;
                  }

                  mockResult = {
                    weather: weather,
                    bgmEnabled: storedBgmEnabled,
                    bgmUrl: bgmUrl,
                    bgm_sun: bgm_sun,
                    bgm_rain: bgm_rain,
                    bgm_snow: bgm_snow,
                    bgm_blossom: bgm_blossom,
                    bgm_leaves: bgm_leaves,
                    realJejuTemp: 18.5,
                    realJejuWind: weather === "blossom" ? 8.5 : 2.1 // 벚꽃 기후일 때 시원한 8.5m/s 강풍 시뮬레이션!
                  };
                } else if (prop === 'updateVillageSettings') {
                  const payload = args[0] || {};
                  localStorage.setItem("mock_weather", payload.weather || "sun");
                  localStorage.setItem("mock_bgm_enabled", String(payload.bgmEnabled || "false"));
                  localStorage.setItem("mock_bgm_url", payload.bgmUrl || "");
                  
                  localStorage.setItem("mock_bgm_sun", payload.bgm_sun || "");
                  localStorage.setItem("mock_bgm_rain", payload.bgm_rain || "");
                  localStorage.setItem("mock_bgm_snow", payload.bgm_snow || "");
                  localStorage.setItem("mock_bgm_blossom", payload.bgm_blossom || "");
                  localStorage.setItem("mock_bgm_leaves", payload.bgm_leaves || "");
                  mockResult = { success: true, message: "[로컬 모크] 마을 기후 마법이 성공적으로 저장되었습니다!" };
                } else if (prop === 'getActiveEvents') {
                  mockResult = [];
                } else if (prop === 'getGeminiApiKey') {
                  const keyVal = localStorage.getItem("mock_gemini_api_key") || "";
                  mockResult = { success: true, key: keyVal };
                } else if (prop === 'setGeminiApiKey') {
                  const newKey = args[0] || "";
                  localStorage.setItem("mock_gemini_api_key", newKey);
                  mockResult = { success: true, message: "[로컬 모크] 구글 스프레드시트 금고에 API Key가 안전하게 저장되었습니다!" };
                } else if (prop === 'submitArchive') {
                  const payload = args[0] || {};
                  let stored = sessionStorage.getItem("mock_feed");
                  let feedList = stored ? JSON.parse(stored) : [];
                  
                  const now = new Date();
                  const dateStr = `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`;
                  const timeStr = `${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;

                  feedList.unshift({
                    name: payload.name || "모험가",
                    item: payload.item || "식단",
                    comment: payload.comment || "",
                    type: payload.type || "식단",
                    photoId: payload.image || "", 
                    date: dateStr,
                    time: timeStr
                  });
                  
                  sessionStorage.setItem("mock_feed", JSON.stringify(feedList));

                  mockResult = {
                    success: true,
                    photoId: "MOCK_PHOTO_ID_" + Date.now(),
                    debugInfo: "MOCK_SAVE_OK"
                  };
                } else if (prop === 'getMyInbodyHistory') {
                  mockResult = {
                    success: true,
                    records: [
                      { date: "2026-05-20", weight: 68.5, muscle: 31.2, fat: 22.4, score: 5.5, memo: "식단 우수 회원! 🌟" },
                      { date: "2026-05-01", weight: 71.0, muscle: 29.5, fat: 25.5, score: 0, memo: "최초 측정일 📊" }
                    ]
                  };
                } else if (prop === 'getUserWellnessActivityHistory') {
                  mockResult = {
                    success: true,
                    todayLogs: [
                      "[회복] 로그인 체크 (+5 점) 🔑",
                      "[회복] 모닝스트레칭 완료 (+2 점) 🤸‍♂️",
                      "[실천] 아침 식단 (Tier S) 인증 (+10 점) 🍱",
                      "[실천] 센터방문 퇴실 (+20 점) 🏛️",
                      "[실천] 운동강도 퇴실 (+80 점) 🤸‍♀️"
                    ],
                    historyLogs: [
                      { date: "26-05-21", score: 117 },
                      { date: "26-05-20", score: 85 },
                      { date: "26-05-19", score: 140 },
                      { date: "26-05-18", score: 30 },
                      { date: "26-05-15", score: 95 },
                      { date: "26-05-14", score: 50 },
                      { date: "26-05-12", score: 110 }
                    ]
                  };
                } else if (prop === 'getVillageNotices') {
                  mockResult = [
                    {
                      date: "2026-05-21",
                      category: "선포",
                      title: "노형 빌리지 대광장 리뉴얼 선포!",
                      content: "노형 빌리지의 광장과 웰니스 센터가 새롭게 태어났습니다. 이제 마을 공지 배너를 누르면 이 황금 두루마리를 통해 언제든 마을의 중대사 역사와 공지 기록을 확인할 수 있습니다. ✨",
                      active: true
                    },
                    {
                      date: "2026-05-20",
                      category: "공지",
                      title: "20시 이후 금식 엄수 권고!",
                      content: "원활한 신체 대사 순환과 다이어트 가속을 위해 밤 20시 이후 밤샘 공복 상태를 완벽히 유지해주시기 바랍니다. 🍵 따뜻한 보리차는 허용됩니다.",
                      active: true
                    }
                  ];
                }
                
                if (sHandler) sHandler(mockResult);
              }, 100);
              return;
            }

            let payload = {};
            if (args.length === 1) {
              if (typeof args[0] === 'object' && args[0] !== null) {
                payload = args[0];
              } else {
                // 단순 값인 경우 args 배열에 담아 보냅니다.
                payload = { args: [args[0]] };
              }
            } else if (args.length > 1) {
              payload = { args: args };
            }

            // Vercel 서버리스 함수(/api/gas)를 통해 통신하여 CORS와 리다이렉트 문제를 해결합니다.
            fetch(`${API_URL}?action=${prop}&t=${Date.now()}`, {
              method: 'POST',
              headers: {
                'Content-Type': 'application/json'
              },
              body: JSON.stringify(payload)
            })
            .then(res => res.json())
            .then(data => {
              if (sHandler) sHandler(data);
            })
            .catch(err => {
              console.error(`[Bridge Error] ${prop}:`, err);
              if (fHandler) fHandler(err);
            });
          };
        }
      })
    }
  };
}

/**
 * [v44.171] 글로벌 인증 및 로그아웃 시스템
 */
const Auth = {
  check: function() {
    const path = window.location.pathname;
    const isLoginPage = path.includes('login.html');
    
    // [v44.207] 인증 결계 제외 목록 (누구나 접근 가능해야 하는 페이지)
    const publicPages = ['attendance.html', 'registration.html', 'login.html'];
    const isPublic = publicPages.some(page => path.includes(page));
    
    if (isLoginPage || isPublic) {
      console.log("🛡️ Auth: Public page or Login page. Skipping check.");
      return;
    }
    
    const name = localStorage.getItem('v44_user_name');
    const phone = localStorage.getItem('v44_user_phone');
    
    console.log(`🛡️ Auth: Checking session... Name: ${name}, Phone: ${phone}`);
    
    if (!name || !phone) {
      console.warn("🛡️ Auth: No session found. Redirecting to login.html...");
      
      // [v44.222] 경로가 루트(/)인 경우를 대비해 절대 경로로 리다이렉트하되, 로컬 file:// 환경에서는 폴더 상대 경로 유지
      let loginUrl;
      if (window.location.protocol === 'file:') {
        const pathParts = window.location.pathname.split('/');
        pathParts[pathParts.length - 1] = 'login.html';
        loginUrl = window.location.protocol + '//' + pathParts.join('/');
      } else {
        loginUrl = window.location.origin + '/login.html';
      }
      
      console.log("🛡️ Auth: Target redirect URL:", loginUrl);
      window.location.replace(loginUrl);
    }
  },
  getPhone: function() {
    return localStorage.getItem('v44_user_phone') || '';
  },
  getName: function() {
    return localStorage.getItem('v44_user_name') || '';
  },
  logout: function() {
    showAppConfirm("🌌 지니 월드에서 퇴장하시겠습니까?", function() {
      localStorage.removeItem('v44_user_name');
      localStorage.removeItem('v44_user_phone');
      localStorage.removeItem('v44_user_tier');
      window.location.replace('login.html');
    }, "🌌");
  }
};

// [v44.207] 모든 페이지 자동 체크 중단 -> 필요한 페이지(index, miracle 등)에서만 명시적으로 호출하도록 변경
// Auth.check(); 

/**
 * [공용] 환경에 맞는 페이지 이동 도우미
 */
function navigateTo(page, params = {}, openInNewTab = false) {
  const query = new URLSearchParams(params).toString();
  const isVercel = window.location.hostname.includes('vercel.app') || window.location.hostname.includes('localhost') || window.location.hostname.includes('127.0.0.1') || window.location.protocol === 'file:';
  
  let url = "";
  if (isVercel) {
    url = page + ".html" + (query ? "?" + query : "");
  } else {
    // GAS 환경
    let baseUrl = (typeof scriptUrl !== 'undefined' && scriptUrl.indexOf('<?') === -1) ? scriptUrl : window.location.href.split('?')[0];
    url = baseUrl + "?page=" + page + (query ? "&" + query : "");
  }

  if (openInNewTab) {
    window.open(url, '_blank');
  } else {
    window.location.href = url;
  }
}


/**
 * [공용] 프리미엄 알림창 (Alert)
 */
function showAppAlert(msg, type = "success", customTitle = "") {
  const modal = document.createElement('div');
  modal.className = 'app-modal-overlay';
  modal.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); backdrop-filter:blur(5px); -webkit-backdrop-filter:blur(5px); display:flex; align-items:center; justify-content:center; z-index:10000; opacity:0; transition:opacity 0.3s; overflow-y:auto; padding:20px 10px; box-sizing:border-box;";
  
  const icon = type === "success" ? "✅" : (type === "error" ? "⚠️" : "ℹ️");
  const title = customTitle || (type === "success" ? "확인 완료" : (type === "error" ? "알림" : "안내"));
  const color = type === "success" ? "#38a169" : (type === "error" ? "#e53e3e" : "#6b46c1");

  const isLargeScreen = window.innerWidth > 700;
  // 단위를 rem 대신 px로 고정하되, 모바일에서 너무 거대하지 않도록 적절히 조절함
  const modalWidth = isLargeScreen ? "650px" : "94%";
  const iconSize = isLargeScreen ? "120px" : "80px";
  const titleSize = isLargeScreen ? "46px" : "30px";
  const textSize = isLargeScreen ? "30px" : "20px";
  const btnSize = isLargeScreen ? "32px" : "22px";
  const btnPadding = isLargeScreen ? "25px" : "18px";

  modal.innerHTML = `
    <div style="background:#fff; width:95%; max-width:${modalWidth}; border-radius:24px; padding:${isLargeScreen ? '40px 30px' : '30px 20px'}; text-align:center; transform:scale(0.8); transition:transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1); box-shadow:0 20px 50px rgba(0,0,0,0.3); border-top: 10px solid ${color};">
      <div style="font-size:${iconSize}; margin-bottom:20px; animation: modalBounce 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);">${icon}</div>
      <h2 style="font-size:${titleSize}; margin-bottom:20px; color:#1a202c; font-weight:900; font-family: sans-serif; letter-spacing:-1px;">${title}</h2>
      <div style="font-size:${textSize}; color:#1a202c; margin-bottom:50px; line-height:1.4; word-break:keep-all; font-family: sans-serif; font-weight:700;">${msg.replace(/\n/g, '<br>')}</div>
      <button onclick="closeAppModal(this)" 
        style="width:100%; padding:${btnPadding}; background:${color}; color:#fff; border:none; border-radius:24px; font-size:${btnSize}; font-weight:900; cursor:pointer; box-shadow: 0 15px 35px rgba(0,0,0,0.2);">
        확 인
      </button>
    </div>
  `;
  document.body.appendChild(modal);
  setTimeout(() => { modal.style.opacity = "1"; modal.firstElementChild.style.transform = "scale(1)"; }, 10);
}

/**
 * [공용] 프리미엄 확인창 (Confirm)
 */
function showAppConfirm(msg, onConfirm, customIcon = "❓", cancelLabel = "취소", confirmLabel = "진행하기") {
  const modal = document.createElement('div');
  modal.className = 'app-modal-overlay';
  modal.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); backdrop-filter:blur(5px); -webkit-backdrop-filter:blur(5px); display:flex; align-items:center; justify-content:center; z-index:10000; opacity:0; transition:opacity 0.3s;";
  
  modal.innerHTML = `
    <div style="background:#fff; width:90%; max-width:380px; border-radius:24px; padding:40px 30px; text-align:center; transform:scale(0.8); transition:transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1); box-shadow:0 20px 40px rgba(0,0,0,0.3); border-top: 6px solid #6b46c1;">
      <div style="font-size:4.5rem; margin-bottom:20px; animation: modalBounce 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);">${customIcon}</div>
      <h2 style="font-size:1.6rem; margin-bottom:12px; color:#1a202c; font-weight:800; font-family: sans-serif;">확인해주세요</h2>
      <p style="font-size:1.05rem; color:#4a5568; margin-bottom:30px; line-height:1.6; word-break:keep-all; font-family: sans-serif;">${msg}</p>
      <div style="display:grid; grid-template-columns: 1fr 1fr; gap:10px; margin-top: 10px;">
        <button onclick="closeAppModal(this)" style="padding:14px 6px; background:#f8fafc; color:#4a5568; border:1px solid #e2e8f0; border-radius:16px; font-size:0.95rem; font-weight:800; cursor:pointer; transition:all 0.2s; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; display:flex; align-items:center; justify-content:center; font-family: sans-serif;" onmouseover="this.style.background='#edf2f7'" onmouseout="this.style.background='#f8fafc'">${cancelLabel}</button>
        <button id="appConfirmOk" style="padding:14px 6px; background:#6b46c1; color:#fff; border:none; border-radius:16px; font-size:0.95rem; font-weight:800; cursor:pointer; box-shadow: 0 6px 15px rgba(107,70,193,0.15); transition:all 0.2s; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; display:flex; align-items:center; justify-content:center; font-family: sans-serif;" onmouseover="this.style.filter='brightness(1.1)'" onmouseout="this.style.filter='none'">${confirmLabel}</button>
      </div>
    </div>
  `;
  document.body.appendChild(modal);
  setTimeout(() => { modal.style.opacity = "1"; modal.firstElementChild.style.transform = "scale(1)"; }, 10);
  
  modal.querySelector('#appConfirmOk').onclick = () => {
    closeAppModal(modal.querySelector('#appConfirmOk'));
    if (onConfirm) onConfirm();
  };
}

function closeAppModal(btn) {
  const modal = btn.closest('.app-modal-overlay');
  modal.style.opacity = "0";
  modal.firstElementChild.style.transform = "scale(0.8)";
  setTimeout(() => modal.remove(), 300);
}

// 애니메이션 스타일 주입
(function injectModalStyles() {
  if (document.getElementById('app-modal-styles')) return;
  const style = document.createElement('style');
  style.id = 'app-modal-styles';
  style.innerHTML = `
    @keyframes modalBounce {
      0% { transform: scale(0.3); opacity: 0; }
      50% { transform: scale(1.1); opacity: 1; }
      70% { transform: scale(0.9); }
      100% { transform: scale(1); }
    }
  `;
  document.head.appendChild(style);
})();

/**
 * [v44.229] Legacy Back Button Trap Removed.
 * Handled by dashboard.js for better UX.
 */


/**
 * [v44.218] 글로벌 로딩 시스템
 */
function showLoading(msg) {
  console.log("⏳ [Silent Sync] " + (msg || "기록을 동기화 중..."));
  // [v50.9] 원장님 지침에 따라 대시보드 및 아카이브 조작을 방해하는 어두운 장막(로딩 마스크)을 완벽하게 영구 무력화합니다!
  const overlay = document.getElementById('v-loading');
  if (overlay) overlay.style.display = 'none';
}
function hideLoading() {
  const overlay = document.getElementById('v-loading');
  if (overlay) overlay.style.display = 'none';
}

/**
 * ==========================================
 * 🌌 [v49.0] 노형 빌리지 글로벌 기후 및 BGM 공유 시스템 (Persistent World Engine)
 * ==========================================
 */
window.initVillageEnvironment = function() {
  // 이미 대시보드 페이지(Village 객체 존재)라면 중복 실행 방지
  if (typeof Village !== 'undefined' || document.getElementById('user-avatar')) {
    console.log("🏰 World Engine: Dashboard page detected. Skipping shared init.");
    return;
  }
  
  if (typeof google === 'undefined' || !google.script || !google.script.run) {
    console.log("🏰 World Engine: Google Apps Script not active yet.");
    return;
  }
  
  google.script.run.withSuccessHandler(function(settings) {
    if (settings) {
      window.applySharedEnvironment(settings);
    }
  }).getVillageSettings();
};

window.applySharedEnvironment = function(settings) {
  if (!settings) return;
  
  const path = window.location.pathname.toLowerCase();
  // [v51.0] 출석체크, 회원가입, 회원재등록, 어드민 등 업무용/키오스크 페이지 및 지니킵 수집기는 날씨 마법사 및 BGM을 원천 차단하여 성능을 보장합니다!
  if (
    path.includes('attendance') || 
    path.includes('registration') || 
    path.includes('renewal') || 
    path.includes('admin') ||
    path.includes('idea_collector')
  ) {
    console.log("🚫 [Shared Environment] Weather and BGM disabled for utility/kiosk page.");
    const oldWrap = document.getElementById('village-weather-wrapper');
    if (oldWrap) oldWrap.remove();
    return;
  }
  
  // [v65.0] 마스터 환경 마법 일괄 제어 토글 처리 (원장님 일괄 차단 지시 수용)
  const magicEnabled = settings.magicEnabled === undefined || settings.magicEnabled.toString().toLowerCase() === 'true';
  if (!magicEnabled) {
    console.log("🚫 [Shared Environment] Village environment magic is globally disabled by the Chief.");
    const oldWrap = document.getElementById('village-weather-wrapper');
    if (oldWrap) oldWrap.remove();
    // BGM 정지 호출
    window.handleSharedBgm(false, '');
    return;
  }
  
  const weatherDisabled = localStorage.getItem('village_weather_disabled') === 'true';
  const weather = weatherDisabled ? 'sun' : (settings.resolvedWeather || settings.weather || 'sun');
  const windSpeed = weatherDisabled ? 0 : (parseFloat(settings.realJejuWind) || 0);
  
  // 1. 기후 파티클 렌더링
  window.renderSharedWeatherParticles(weather, windSpeed);
  
  // 2. BGM 재생 및 토글 제어
  const bgmEnabled = settings.bgmEnabled && settings.bgmEnabled.toString().toLowerCase() === 'true';
  const bgmUrl = settings.bgmUrl || '';
  window.handleSharedBgm(bgmEnabled, bgmUrl);
};

window.renderSharedWeatherParticles = function(weather, windSpeed = 0) {
  const oldWrap = document.getElementById('village-weather-wrapper');
  if (oldWrap) oldWrap.remove();
  
  if (localStorage.getItem('village_weather_disabled') === 'true') {
    return;
  }
  
  const isWindy = windSpeed >= 5.0;
  
  const wrapper = document.createElement('div');
  wrapper.id = 'village-weather-wrapper';
  wrapper.style = "position:fixed; top:0; left:0; width:100%; height:100%; pointer-events:none; z-index:9999; overflow:hidden;";
  document.body.appendChild(wrapper);
  
  if (isWindy) {
    const windWispCount = 12;
    for (let i = 0; i < windWispCount; i++) {
      const wisp = document.createElement('div');
      wisp.style.position = 'absolute';
      wisp.style.top = (Math.random() * 90) + 'vh';
      wisp.style.left = '-250px';
      wisp.style.width = (Math.random() * 180 + 100) + 'px';
      wisp.style.height = '1.2px';
      wisp.style.background = 'linear-gradient(90deg, transparent, rgba(255, 255, 255, 0.28), transparent)';
      wisp.style.opacity = Math.random() * 0.5 + 0.15;
      
      const duration = Math.max(1.2, 5.5 - (windSpeed * 0.35)); 
      const delay = Math.random() * 5;
      wisp.style.animation = `blow-wind ${duration}s linear ${delay}s infinite`;
      wrapper.appendChild(wisp);
    }
  }
  
  // 📢 맑은 날(sun)의 몽환적이고 고선명 무지갯빛 비눗방울 효과 추가 (밝은 화면에서도 선명히 보이도록 최적화)
  if (weather === 'sun') {
    const bubbleCount = 15;
    for (let i = 0; i < bubbleCount; i++) {
      const b = document.createElement('div');
      b.style.position = 'absolute';
      b.style.bottom = '-35px';
      b.style.left = Math.random() * 100 + 'vw';
      
      // 크기를 조금 더 다양하게 (6px ~ 22px) 확대하여 시인성 보장
      const size = Math.random() * 16 + 6; 
      b.style.width = size + 'px';
      b.style.height = size + 'px';
      b.style.borderRadius = '50%';
      
      // 🌈 밝은 화면에서도 빛나는 환상적인 무지갯빛 그라데이션 (Pastel Rainbow Refraction)
      b.style.background = 'radial-gradient(circle at 35% 35%, rgba(255, 255, 255, 0.9) 0%, rgba(255, 220, 230, 0.5) 20%, rgba(220, 240, 255, 0.45) 50%, rgba(220, 255, 220, 0.4) 75%, rgba(255, 235, 190, 0.55) 100%)';
      
      // 테두리를 더 뚜렷하고 입체적인 반투명 화이트로 변경
      b.style.border = '0.8px solid rgba(255, 255, 255, 0.75)';
      
      // 무지갯빛 반사광 효과를 배가해주는 내부 이중 음영 및 은은한 외곽 글로우 추가
      b.style.boxShadow = 'inset -2px -2px 5px rgba(135, 206, 250, 0.45), inset 2px 2px 5px rgba(255, 182, 193, 0.55), 0 3px 6px rgba(165, 94, 234, 0.12)';
      
      // 밝은 화면에서도 보일 수 있게 기본 불투명도 범위를 대폭 상향조정 (0.35 ~ 0.75)
      b.style.opacity = Math.random() * 0.4 + 0.35; 
      b.style.pointerEvents = 'none';
      
      // 🐢 훨씬 천천히 아지랑이처럼 동실동실 떠오르도록 애니메이션 재생 시간 대폭 연장 (14초 ~ 24초)
      const duration = Math.random() * 10 + 14; 
      const delay = Math.random() * 12;
      
      const animName = `rise-bubble-${i % 3}`;
      b.style.animation = `${animName} ${duration}s ease-in-out ${delay}s infinite`;
      wrapper.appendChild(b);
    }
  }
  
  if (weather !== 'sun') {
    const particleCount = weather === 'rain' ? 80 : weather === 'snow' ? 50 : 32;
    const chars = {
      snow: '❄',
      blossom: '🌸',
      leaves: '🍁'
    };
    
    for (let i = 0; i < particleCount; i++) {
      const p = document.createElement('div');
      p.style.position = 'absolute';
      p.style.top = '-20px';
      p.style.left = Math.random() * 100 + 'vw';
      p.style.opacity = Math.random() * 0.7 + 0.3;
      
      const size = Math.random() * 15 + 10;
      p.style.fontSize = size + 'px';
      
      let duration = Math.random() * 5 + 5;
      if (isWindy) duration = duration * 0.45;
      
      const delay = Math.random() * 8;
      
      if (weather === 'rain') {
        p.style.width = '1.5px';
        p.style.height = (Math.random() * 20 + 15) + 'px';
        p.style.background = 'rgba(174, 219, 255, 0.6)';
        const angle = isWindy ? 35 : 15;
        p.style.transform = `rotate(${angle}deg)`;
      } else {
        p.innerText = chars[weather] || '';
      }
      
      const animName = isWindy ? `fall-${weather}-windy` : `fall-${weather}`;
      p.style.animation = `${animName} ${duration}s linear ${delay}s infinite`;
      wrapper.appendChild(p);
    }
  }
  
  const styleId = 'weather-keyframes-style';
  if (!document.getElementById(styleId)) {
    const style = document.createElement('style');
    style.id = styleId;
    style.innerHTML = `
      @keyframes blow-wind { 0% { transform: translateX(-250px); } 100% { transform: translateX(110vw); } }
      @keyframes fall-rain { to { transform: translateY(105vh) rotate(15deg); } }
      @keyframes fall-rain-windy { to { transform: translateY(105vh) translateX(35vw) rotate(35deg); } }
      @keyframes fall-snow { 0% { transform: translateY(-20px) translateX(0) rotate(0deg); } 50% { transform: translateY(50vh) translateX(20px) rotate(180deg); } 100% { transform: translateY(105vh) translateX(-10px) rotate(360deg); } }
      @keyframes fall-snow-windy { 0% { transform: translateY(-20px) translateX(0) rotate(0deg); } 100% { transform: translateY(105vh) translateX(65vw) rotate(720deg); } }
      @keyframes fall-blossom { 0% { transform: translateY(-20px) translateX(0) rotate(0deg); } 50% { transform: translateY(50vh) translateX(30px) rotate(120deg); } 100% { transform: translateY(105vh) translateX(-20px) rotate(240deg); } }
      @keyframes fall-blossom-windy { 0% { transform: translateY(-20px) translateX(0) rotate(0deg); } 100% { transform: translateY(105vh) translateX(75vw) rotate(540deg); } }
      @keyframes fall-leaves { 0% { transform: translateY(-20px) translateX(0) rotate(0deg); } 50% { transform: translateY(50vh) translateX(-20px) rotate(180deg); } 100% { transform: translateY(105vh) translateX(15px) rotate(360deg); } }
      @keyframes fall-leaves-windy { 0% { transform: translateY(-20px) translateX(0) rotate(0deg); } 100% { transform: translateY(105vh) translateX(70vw) rotate(480deg); } }
      @keyframes rise-bubble-0 {
        0% { transform: translateY(0) translateX(0); opacity: 0; }
        10% { opacity: 0.5; }
        50% { transform: translateY(-50vh) translateX(15px); }
        90% { opacity: 0.5; }
        100% { transform: translateY(-110vh) translateX(-10px); opacity: 0; }
      }
      @keyframes rise-bubble-1 {
        0% { transform: translateY(0) translateX(0); opacity: 0; }
        15% { opacity: 0.4; }
        50% { transform: translateY(-50vh) translateX(-20px); }
        85% { opacity: 0.4; }
        100% { transform: translateY(-110vh) translateX(15px); opacity: 0; }
      }
      @keyframes rise-bubble-2 {
        0% { transform: translateY(0) translateX(0); opacity: 0; }
        10% { opacity: 0.6; }
        50% { transform: translateY(-50vh) translateX(10px); }
        90% { opacity: 0.6; }
        100% { transform: translateY(-110vh) translateX(-15px); opacity: 0; }
      }
    `;
    document.head.appendChild(style);
  }
};

window.handleSharedBgm = function(enabled, url) {
  if (!url) return;
  
  // 쉼표(,), 세미콜론(;), 줄바꿈/개행(\n, \r) 기준으로 쪼개어 플레이리스트 생성
  const playlist = url.split(/[,;\n\r]+/).map(u => u.trim()).filter(u => u.length > 0);
  if (playlist.length === 0) return;
  
  let audio = document.getElementById('village-bgm-audio');
  let isNewPlaylist = false;
  
  if (!audio) {
    audio = document.createElement('audio');
    audio.id = 'village-bgm-audio';
    document.body.appendChild(audio);
    isNewPlaylist = true;
  } else {
    // [호환성 방어] dataset 대신 getAttribute/setAttribute를 활용하여 구형 기기 크래시 방지
    const oldPlaylistStr = audio.getAttribute('data-playlist') || "";
    const newPlaylistStr = playlist.join(",");
    if (oldPlaylistStr !== newPlaylistStr) {
      isNewPlaylist = true;
    }
  }
  
  if (isNewPlaylist) {
    audio.setAttribute('data-playlist', playlist.join(","));
    audio.setAttribute('data-current-index', "0");
    audio.src = playlist[0];
    
    // 여러 곡일 경우 ended 이벤트를 이용하여 순환, 단일 곡일 경우 loop 처리
    if (playlist.length > 1) {
      audio.loop = false;
      audio.onended = function() {
        let currentIndex = parseInt(audio.getAttribute('data-current-index') || "0", 10);
        let nextIndex = (currentIndex + 1) % playlist.length;
        audio.setAttribute('data-current-index', String(nextIndex));
        audio.src = playlist[nextIndex];
        
        // [호환성 방어] play() 반환 Promise가 undefined일 수 있는 구형 모바일 브라우저 예외 처리
        const playPromise = audio.play();
        if (playPromise !== undefined) {
          playPromise.then(function() {
            console.log("Playing next shared BGM track:", playlist[nextIndex]);
          }).catch(function(e) { console.log("Next track play failed:", e); });
        }
      };
    } else {
      audio.loop = true;
      audio.onended = null;
    }
  }
  
  let btn = document.getElementById('village-bgm-toggle');
  if (!enabled) {
    if (btn) btn.style.display = 'none';
    audio.pause();
    return;
  }
  
  if (btn) {
    btn.style.display = 'flex';
  } else {
    btn = document.createElement('div');
    btn.id = 'village-bgm-toggle';
    btn.style = "position:fixed; bottom:90px; right:15px; width:45px; height:45px; border-radius:50%; background:rgba(255,255,255,0.85); box-shadow:0 8px 32px rgba(31,38,135,0.15); backdrop-filter:blur(6px); border:1.5px solid rgba(255,255,255,0.18); display:flex; align-items:center; justify-content:center; cursor:pointer; z-index:9998; transition:transform 0.3s ease;";
    btn.innerHTML = `<i class="fas fa-music" style="color:var(--v-wood); font-size:1.1rem; transition: color 0.3s;"></i>`;
    document.body.appendChild(btn);
    
    btn.onclick = function() {
      if (audio.paused) {
        const playPromise = audio.play();
        if (playPromise !== undefined) {
          playPromise.then(function() {
            btn.querySelector('i').style.color = '#e74c3c';
            btn.style.transform = 'scale(1.1) rotate(15deg)';
            localStorage.setItem('v44_bgm_user_play', 'true');
          }).catch(function(e) { console.log("BGM Play error:", e); });
        }
      } else {
        audio.pause();
        btn.querySelector('i').style.color = 'var(--v-wood)';
        btn.style.transform = 'scale(1) rotate(0deg)';
        localStorage.setItem('v44_bgm_user_play', 'false');
      }
    };
  }
  
  const userPlayPreference = localStorage.getItem('v44_bgm_user_play');
  const shouldPlay = userPlayPreference !== 'false';
  
  if (shouldPlay) {
    const startPlay = function() {
      const playPromise = audio.play();
      if (playPromise !== undefined) {
        playPromise.then(function() {
          btn.querySelector('i').style.color = '#e74c3c';
          btn.style.transform = 'scale(1.1) rotate(15deg)';
          document.removeEventListener('click', startPlay);
          document.removeEventListener('touchstart', startPlay);
        }).catch(function(e) { console.log("BGM play failed on interaction:", e); });
      }
    };
    
    const playPromise = audio.play();
    if (playPromise !== undefined) {
      playPromise.then(function() {
        btn.querySelector('i').style.color = '#e74c3c';
        btn.style.transform = 'scale(1.1) rotate(15deg)';
      }).catch(function(e) {
        console.log("Auto-play blocked by browser; registering interaction fallback.");
        document.addEventListener('click', startPlay);
        document.addEventListener('touchstart', startPlay);
      });
    }
  } else {
    audio.pause();
    btn.querySelector('i').style.color = 'var(--v-wood)';
    btn.style.transform = 'scale(1) rotate(0deg)';
  }
};

// 모든 공용 페이지 로드 시 기후 및 BGM 자동 기동
window.addEventListener('load', function() {
  if (window.initVillageEnvironment) {
    window.initVillageEnvironment();
  }
});

// ----------------------------------------------------
// 🌌 [Genie World] 글로벌 모바일 물리 뒤로가기 하드웨어 백버튼 트랩 시스템
// ----------------------------------------------------
window.globalModalHistoryStack = [];

window.registerModalOpen = function(closeFn) {
  if (typeof closeFn !== 'function') return;
  // 가상 상태를 쌓아 뒤로가기 흐름을 감지함
  history.pushState({ modalOpen: true }, null, location.href);
  window.globalModalHistoryStack.push(closeFn);
  console.log("[Global Backbutton] Modal registered. Stack size:", window.globalModalHistoryStack.length);
};

window.registerModalClose = function(closeFn) {
  const idx = window.globalModalHistoryStack.indexOf(closeFn);
  if (idx > -1) {
    window.globalModalHistoryStack.splice(idx, 1);
    window._isManualModalClose = true; // [v50.9.1] 수동 모달 닫기 플래그 활성화 (popstate 퇴장 확인팝업 방어용)
    history.back();
    console.log("[Global Backbutton] Modal deregistered. Stack size:", window.globalModalHistoryStack.length);
  }
};

window.addEventListener('popstate', function(e) {
  if (window._isManualModalClose) {
    window._isManualModalClose = false; // 플래그 즉시 리셋
    console.log("[Global Backbutton] Ignored popstate exit confirmation because of manual modal close.");
    return;
  }

  if (window.globalModalHistoryStack.length > 0) {
    // 뒤로가기 방어: 스택에서 가장 최근 등록된 닫기 콜백을 실행하여 모달을 닫아줍니다!
    const closeFn = window.globalModalHistoryStack.pop();
    if (closeFn) {
      closeFn(true); // popstate를 통한 물리적 닫힘으로 호출
    }
  } else {
    const path = window.location.pathname;
    const pageName = path.substring(path.lastIndexOf('/') + 1) || 'index.html';

    // 💡 [원상 복귀] 테라피 예약 화면(reservation.html)은 원래대로 로컬 popstate가 통제하도록 글로벌 가드 bypass!
    if (pageName === 'reservation.html') {
      return;
    }

    // 열려있는 모달이 하나도 없을 때, 현재 화면에 알맞게 안전 퇴장 컨펌 작동!
    history.pushState(null, null, location.href); // 히스토리 밀림 복구
    
    if (pageName === 'admin.html') {
      showAppConfirm("👑 점핑관리자 앱을 종료하시겠습니까?", function() {
        localStorage.removeItem('v44_user_name');
        localStorage.removeItem('v44_user_phone');
        window.close();
        setTimeout(function() { history.go(-2); }, 50);
      }, "👑", "아니오(머무르기)", "예(앱종료)");
    } else if (pageName === 'attendance.html') {
      showAppConfirm("⚙️ 출석체크 앱을 종료하시겠습니까?", function() {
        window.close();
        setTimeout(function() { history.go(-2); }, 50);
      }, "⚙️", "아니오(머무르기)", "예(앱종료)");
    } else if (pageName === 'registration.html') {
      showAppConfirm("⚙️ 회원등록앱을 종료하시겠습니까?", function() {
        window.close();
        setTimeout(function() { history.go(-2); }, 50);
      }, "⚙️", "아니오(머무르기)", "예(저장취소)");
    } else if (pageName === 'renewal.html' || pageName.indexOf('renewal') !== -1) {
      showAppConfirm("⚙️ 회원등록앱을 종료하시겠습니까?", function() {
        window.close();
        setTimeout(function() { history.go(-2); }, 50);
      }, "⚙️", "아니오(머무르기)", "예(저장취소)");
    } else if (pageName === 'oasis.html') {
      showAppConfirm("🌴 오아시스에서 퇴장하여 메인 광장으로 이동하시겠습니까?", function() {
        window.location.replace('index.html');
      }, "🌴");
    } else if (pageName === 'miracle.html') {
      showAppConfirm("⚔️ 모험가의 기록소에서 퇴장하여 메인 광장으로 이동하시겠습니까?", function() {
        window.location.replace('index.html');
      }, "⚔️");
    } else if (pageName === 'wisdom.html') {
      showAppConfirm("📜 지식창고에서 퇴장하여 메인 광장으로 이동하시겠습니까?", function() {
        window.location.replace('index.html');
      }, "📜");
    } else if (pageName === 'index.html' || pageName === '') {
      showAppConfirm("🏰 지니 월드(노형 빌리지)에서 완전히 퇴장하시겠습니까?", function() {
        const logoutBtn = document.getElementById('btn-logout');
        if (logoutBtn) {
          logoutBtn.click();
        } else {
          alert("우측 상단의 퇴장/로그아웃 버튼을 눌러 안전하게 퇴장해 주세요! ✨");
        }
      }, "🏰");
    }
  }
});

// 진입 시 기본 히스토리 백 트랩 시드 생성
window.addEventListener('DOMContentLoaded', function() {
  history.pushState(null, null, location.href);
});

/**
 * ==========================================
 * 📬 [v59.0] 지니의 마을 우편함 & 전령 포털 엔진 (글로벌 렌더링 & 연동)
 * ==========================================
 */
const Mailbox = {
  notifications: [],
  initialized: false,
  markingAsRead: false,
  
  init: function() {
    if (this.initialized) return;
    
    // 제외 페이지 필터링
    const path = window.location.pathname.toLowerCase();
    if (
      path.includes('attendance') || 
      path.includes('registration') || 
      path.includes('renewal') || 
      path.includes('admin') ||
      path.includes('login') ||
      path.includes('idea_collector')
    ) {
      console.log("🚫 [Mailbox] Disabled for utility/kiosk page.");
      return;
    }
    
    const phone = Auth.getPhone();
    if (!phone) {
      console.log("🚫 [Mailbox] No session found. Waiting for login.");
      return;
    }
    
    this.initialized = true;
    this.injectStyles();
    this.injectMarkup();
    this.loadNotifications(phone);
  },
  
  injectStyles: function() {
    const styleId = 'mailbox-global-styles';
    if (document.getElementById(styleId)) return;
    
    const style = document.createElement('style');
    style.id = styleId;
    style.innerHTML = `
      /* 우편함 FAB 버튼 */
      .v-mailbox-fab {
        position: fixed;
        bottom: 90px;
        left: 15px;
        width: 45px;
        height: 45px;
        border-radius: 50%;
        background: rgba(255, 255, 255, 0.88);
        box-shadow: 0 8px 32px rgba(31, 38, 135, 0.18);
        backdrop-filter: blur(6px);
        -webkit-backdrop-filter: blur(6px);
        border: 1.5px solid rgba(255, 255, 255, 0.25);
        display: flex;
        align-items: center;
        justify-content: center;
        cursor: pointer;
        z-index: 9998;
        transition: all 0.3s cubic-bezier(0.34, 1.56, 0.64, 1);
        animation: mailbox-breathing 3.5s ease-in-out infinite;
      }
      .v-mailbox-fab:hover {
        transform: scale(1.1);
        background: #ffffff;
      }
      .v-mailbox-fab i {
        color: #4f46e5;
        font-size: 1.15rem;
      }
      
      /* 우편함 알림 배지 */
      .mailbox-badge {
        position: absolute;
        top: -2px;
        right: -2px;
        width: 13px;
        height: 13px;
        border-radius: 50%;
        background: #ff4757;
        border: 2px solid #ffffff;
        box-shadow: 0 0 10px #ff4757;
        animation: pulse-unread 1.5s infinite ease-in-out;
      }
      
      /* 애니메이션 정의 */
      @keyframes mailbox-breathing {
        0% { transform: translateY(0); }
        50% { transform: translateY(-7px); }
        100% { transform: translateY(0); }
      }
      @keyframes pulse-unread {
        0% { transform: scale(0.9); opacity: 0.5; box-shadow: 0 0 4px #ff4757; }
        50% { transform: scale(1.1); opacity: 1; box-shadow: 0 0 12px #ff4757; }
        100% { transform: scale(0.9); opacity: 0.5; box-shadow: 0 0 4px #ff4757; }
      }
      @keyframes pulse-unread-fab {
        0% { box-shadow: 0 8px 32px rgba(31, 38, 135, 0.18); }
        50% { box-shadow: 0 8px 32px rgba(239, 68, 68, 0.25), 0 0 15px rgba(239, 68, 68, 0.35); border-color: rgba(239, 68, 68, 0.4); }
        100% { box-shadow: 0 8px 32px rgba(31, 38, 135, 0.18); }
      }
      
      /* 우편함 모달 & 리스트 */
      .mailbox-card {
        background: #ffffff;
        border-radius: 18px;
        padding: 14px 16px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.03);
        border: 1px solid #edf2f7;
        display: flex;
        flex-direction: column;
        gap: 6px;
        transition: all 0.25s ease;
        text-align: left;
        position: relative;
        overflow: hidden;
        flex-shrink: 0; /* 우편함 개수 증가 시 카드 찌그러짐 방지 */
      }
      .mailbox-card.unread {
        box-shadow: 0 6px 20px rgba(79, 70, 229, 0.08);
      }
      .mailbox-card.read {
        opacity: 0.65;
        filter: grayscale(20%);
        background: #f8fafc;
      }
      .mailbox-card .card-header {
        display: flex;
        justify-content: space-between;
        align-items: flex-start; /* 타이틀 여러 줄일 때 정렬 깨짐 방지 */
        gap: 12px;
      }
      .mailbox-card .card-title {
        font-size: 0.85rem;
        font-weight: 800;
        color: #1e293b;
        display: flex;
        align-items: flex-start; /* 이모지 높이 맞춤 */
        gap: 6px;
        flex: 1;
        word-break: keep-all;
        overflow-wrap: anywhere;
        line-height: 1.4;
      }
      .mailbox-card .card-time {
        font-size: 0.65rem;
        font-weight: 600;
        color: #94a3b8;
        white-space: nowrap; /* 시간 텍스트 줄바꿈 깨짐 완전 차단 */
        flex-shrink: 0;
        margin-top: 2px;
      }
      .mailbox-card .card-body {
        font-size: 0.73rem;
        font-weight: 700;
        color: #475569;
        line-height: 1.45;
        word-break: keep-all;
        overflow-wrap: break-word;
      }
      
      /* 우편함 스크롤바 모바일/크롬 디자인 튜닝 */
      #mailbox-list::-webkit-scrollbar {
        width: 5px;
      }
      #mailbox-list::-webkit-scrollbar-track {
        background: transparent;
      }
      #mailbox-list::-webkit-scrollbar-thumb {
        background: #cbd5e1;
        border-radius: 10px;
      }
      #mailbox-list::-webkit-scrollbar-thumb:hover {
        background: #94a3b8;
      }
      
      /* 유형별 특수 보더 및 배경 효과 */
      .mailbox-card.type-welcome {
        border-left: 5px solid #10b981;
      }
      .mailbox-card.type-debuff {
        border-left: 5px solid #ef4444;
        animation: card-warning-pulse 2.2s infinite ease-in-out;
      }
      .mailbox-card.type-quest {
        border-left: 5px solid #ec4899;
      }
      .mailbox-card.type-admin {
        border-left: 5px solid #3b82f6;
      }
      .mailbox-card.type-ranking {
        border-left: 5px solid #8a2be2;
        background: linear-gradient(135deg, #ffffff 0%, rgba(138, 43, 226, 0.03) 100%);
      }
      
      .mailbox-card.read.type-welcome,
      .mailbox-card.read.type-debuff,
      .mailbox-card.read.type-quest,
      .mailbox-card.read.type-admin,
      .mailbox-card.read.type-ranking {
        border-left-color: #cbd5e1;
        background: #f8fafc;
      }
      
      /* [v60.0] 명예의 전당 액션 버튼 스타일 */
      .mailbox-action-btn {
        width: 100%;
        background: linear-gradient(135deg, #8a2be2, #4f46e5);
        color: #ffffff;
        border: none;
        border-radius: 12px;
        padding: 10px 14px;
        font-size: 0.73rem;
        font-weight: 800;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 6px;
        box-shadow: 0 4px 12px rgba(138, 43, 226, 0.25);
        transition: all 0.2s ease;
        margin-top: 6px;
        box-sizing: border-box;
      }
      .mailbox-action-btn:hover {
        transform: translateY(-1.5px);
        box-shadow: 0 6px 16px rgba(138, 43, 226, 0.35);
        filter: brightness(1.05);
      }
      .mailbox-action-btn:active {
        transform: translateY(0);
      }
      
      @keyframes card-warning-pulse {
        0% { border-left-color: #ef4444; box-shadow: 0 4px 15px rgba(239, 68, 68, 0.03); }
        50% { border-left-color: #f87171; box-shadow: 0 4px 20px rgba(239, 68, 68, 0.08); }
        100% { border-left-color: #ef4444; box-shadow: 0 4px 15px rgba(239, 68, 68, 0.03); }
      }
      
      .new-badge {
        background: #ef4444;
        color: #ffffff;
        font-size: 0.55rem;
        font-weight: 900;
        padding: 1.5px 5px;
        border-radius: 6px;
        text-transform: uppercase;
        display: inline-block;
      }
    `;
    document.head.appendChild(style);
  },
  
  injectMarkup: function() {
    // 1. FAB 생성
    if (!document.getElementById('mailbox-fab')) {
      const fab = document.createElement('div');
      fab.id = 'mailbox-fab';
      fab.className = 'v-mailbox-fab';
      fab.innerHTML = `
        <i class="fa-solid fa-envelope"></i>
        <span class="mailbox-badge" id="mailbox-unread-badge" style="display: none;"></span>
      `;
      fab.onclick = () => this.open();
      document.body.appendChild(fab);
    }
    
    // 2. 모달 생성
    if (!document.getElementById('mailbox-modal')) {
      const modal = document.createElement('div');
      modal.id = 'mailbox-modal';
      modal.style = "display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(15,23,42,0.65); backdrop-filter:blur(12px); -webkit-backdrop-filter:blur(12px); z-index:99999; align-items:flex-end; justify-content:center;";
      
      modal.innerHTML = `
        <div style="background:#fff; border-radius:30px 30px 0 0; width:100%; max-width:500px; padding:25px 20px 45px; box-shadow:0 -8px 40px rgba(0,0,0,0.15); animation:slideUp 0.35s ease-out; text-align: left; box-sizing: border-box; position: relative;">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:18px;">
            <div style="font-size:1.05rem; font-weight:900; color:#4f46e5; display:flex; align-items:center; gap:6px;">
              ✉️ 지니의 마을 우편함 <span id="mailbox-count-text" style="font-size:0.68rem; font-weight:800; color:#64748b; background:#f1f5f9; padding:2px 8px; border-radius:10px; display:inline-block;">새 쪽지 0개</span>
            </div>
            <button onclick="Mailbox.close()" style="background:none; border:none; font-size:1.3rem; color:#94a3b8; cursor:pointer; font-weight:bold; font-family:inherit;">✕</button>
          </div>
          
          <!-- 리스트 -->
          <div id="mailbox-list" style="max-height: 380px; overflow-y: auto; padding-right: 4px; box-sizing: border-box; display: flex; flex-direction: column; gap: 10px;">
            <!-- 동적 삽입 -->
          </div>
          
          <!-- 텅 빈 상태 -->
          <div id="mailbox-empty-view" style="display:none; text-align:center; padding:50px 20px; color:#94a3b8;">
            <span style="font-size: 2.8rem; display:block; margin-bottom:12px;">✉️</span>
            <div style="font-size: 0.8rem; font-weight: 900; color:#475569;">우편함이 텅 비어 있습니다.</div>
            <div style="font-size: 0.65rem; font-weight: 700; color:#94a3b8; margin-top:3px;">안부 편지나 활동 알림이 오면 여기에 쌓입니다.</div>
          </div>
        </div>
      `;
      document.body.appendChild(modal);
    }
  },
  
  loadNotifications: function(phone) {
    if (typeof google === 'undefined' || !google.script || !google.script.run) {
      // 로컬 Mock 모드 또는 오프라인 테스트용 데이터 바인딩
      console.log("🎮 [Mailbox] Mock data loading...");
      this.notifications = [
        { id: "MOCK_1", type: "welcome", title: "오늘 첫 로그인! 웰니스 보너스 지급 완료", content: "점핑님! 오늘 하루도 힘차게 시작해봐요! 로그인 보너스로 +5 EXP(회복력)가 즉시 지급되었습니다. ⚔️", createdAt: "2026-05-22 08:00:00", isRead: false },
        { id: "MOCK_2", type: "debuff", title: "연속 미출석 에너지 방전 디버프 발생!", content: "김점핑 회원님! 클럽 출석을 하지 않으신 지 연속 4일이 경과하여 에너지가 방전되었습니다. 누적 점수에서 -100 EXP가 차감되었습니다. 오늘 출석 즉시 100% 자동 부활합니다!", createdAt: "2026-05-21 07:30:00", isRead: true }
      ];
      this.render();
      return;
    }
    
    const self = this;
    google.script.run.withSuccessHandler(function(res) {
      if (res && res.success) {
        self.notifications = res.notifications || [];
        self.render();
      }
    }).getPersonalNotifications({ phone: phone });
  },
  
  render: function() {
    const listContainer = document.getElementById('mailbox-list');
    const emptyView = document.getElementById('mailbox-empty-view');
    const badge = document.getElementById('mailbox-unread-badge');
    const countText = document.getElementById('mailbox-count-text');
    const fab = document.getElementById('mailbox-fab');
    
    if (!listContainer) return;
    
    listContainer.innerHTML = "";
    
    const unreadCount = this.notifications.filter(n => !n.isRead).length;
    
    // FAB 알림 배지 제어 및 흔들림(breathing) 애니메이션 극대화
    if (unreadCount > 0) {
      if (badge) badge.style.display = "block";
      if (fab) {
        fab.style.animation = "mailbox-breathing 1.8s ease-in-out infinite, pulse-unread-fab 1.8s infinite ease-in-out";
      }
    } else {
      if (badge) badge.style.display = "none";
      if (fab) {
        fab.style.animation = "mailbox-breathing 3.5s ease-in-out infinite"; // 차분한 호버
      }
    }
    
    if (countText) {
      countText.innerText = `새 쪽지 ${unreadCount}개`;
      countText.style.background = unreadCount > 0 ? "rgba(239, 68, 68, 0.08)" : "#f1f5f9";
      countText.style.color = unreadCount > 0 ? "#ef4444" : "#64748b";
    }
    
    if (this.notifications.length === 0) {
      listContainer.style.display = "none";
      if (emptyView) emptyView.style.display = "block";
      return;
    }
    
    listContainer.style.display = "flex";
    if (emptyView) emptyView.style.display = "none";
    
    this.notifications.forEach(noti => {
      const card = document.createElement('div');
      card.className = `mailbox-card ${noti.isRead ? 'read' : 'unread'} type-${noti.type}`;
      
      const emojiMap = { 
        welcome: "🎉", 
        "웰컴": "🎉",
        debuff: "⚠️", 
        "디버프": "⚠️",
        quest: "⚡", 
        "퀘스트": "⚡",
        "방어": "🛡️",
        "안부": "💌",
        admin: "✉️",
        ranking: "🏆",
        "랭킹": "🏆"
      };
      const emoji = emojiMap[noti.type] || "✉️";
      
      const timeStr = this.timeAgo(noti.createdAt);
      
      // [v60.0] ranking 유형일 때 하단 바로가기 액션 버튼 삽입
      const actionButtonHtml = noti.type === 'ranking' 
        ? `<button onclick="navigateTo('halloffame')" class="mailbox-action-btn">🏆 명예의 전당 바로가기 ➔</button>` 
        : '';
      
      card.innerHTML = `
        <div class="card-header">
          <div class="card-title">
            <span>${emoji}</span>
            <span>${noti.title}</span>
            ${!noti.isRead ? '<span class="new-badge">NEW</span>' : ''}
          </div>
          <span class="card-time">${timeStr}</span>
        </div>
        <div class="card-body">
          ${noti.content.replace(/\n/g, '<br>')}
        </div>
        ${actionButtonHtml}
      `;
      listContainer.appendChild(card);
    });
  },
  
  open: function() {
    const modal = document.getElementById('mailbox-modal');
    if (!modal) return;
    
    modal.style.display = "flex";
    
    // 뒤로가기 하드웨어 버튼 연동 등록
    if (window.registerModalOpen) {
      window.registerModalOpen((isBack) => {
        modal.style.display = "none";
      });
    }
    
    // 열람했으므로 서버 읽음 처리 API 비동기 격발
    const phone = Auth.getPhone();
    if (!phone) return;
    
    const self = this;
    const hasUnread = this.notifications.some(n => !n.isRead);
    if (hasUnread && !this.markingAsRead) {
      this.markingAsRead = true;
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        google.script.run
          .withSuccessHandler(function(res) {
            self.markingAsRead = false;
            if (res && res.success) {
              // 로컬 상태 즉각 전부 읽음 처리
              self.notifications.forEach(n => n.isRead = true);
              self.render();
            }
          })
          .withFailureHandler(function(err) {
            self.markingAsRead = false;
          })
          .markNotificationsAsRead({ phone: phone });
      } else {
        // Mock 모드 대응
        setTimeout(() => {
          self.notifications.forEach(n => n.isRead = true);
          self.render();
          self.markingAsRead = false;
        }, 300);
      }
    }
  },
  
  close: function() {
    const modal = document.getElementById('mailbox-modal');
    if (modal) {
      modal.style.display = "none";
      if (window.registerModalClose) {
        window.registerModalClose((isBack) => {
          modal.style.display = "none";
        });
      }
    }
  },
  
  timeAgo: function(dateStr) {
    if (!dateStr) return "";
    try {
      var parts = dateStr.split(/[ \-:]/);
      if (parts.length < 6) return dateStr;
      
      var createdDate = new Date(parts[0], parts[1] - 1, parts[2], parts[3], parts[4], parts[5]);
      
      // [v59.1] 구글 스프레드시트 빈 값/0 입력 시 1899~1900년으로 자동 파싱되어 '46194일전' 등으로 노출되는 현상 방어
      if (createdDate.getFullYear() < 1905) {
        return "알 수 없는 날짜";
      }
      
      var now = new Date();
      var diffTime = now.getTime() - createdDate.getTime();
      
      var diffMinutes = Math.floor(diffTime / (1000 * 60));
      if (diffMinutes < 1) return "방금 전";
      if (diffMinutes < 60) return diffMinutes + "분 전";
      
      var diffHours = Math.floor(diffMinutes / 60);
      if (diffHours < 24) return diffHours + "시간 전";
      
      var diffDays = Math.floor(diffHours / 24);
      if (diffDays === 1) return "어제";
      return diffDays + "일 전";
    } catch (e) {
      return dateStr;
    }
  }
};

// 대시보드 로드 데이터 자동 동기화용 리스너 등록
window.syncMailboxWithDashboardData = function(res) {
  if (res && res.success && res.notifications) {
    Mailbox.notifications = res.notifications;
    Mailbox.render();
  }
};

// 로드 시 우편함 기동
window.addEventListener('load', function() {
  Mailbox.init();
});
