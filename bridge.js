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
                    pillarNotice: {
                      content: "노형 빌리지 리뉴얼 완공! 🏰 따사로운 봄바람과 함께 벚꽃 마법이 시작되었습니다. 🌸 우측 하단의 BGM 플레이어를 켜고 낭만 넘치는 음악을 감상해 보세요!"
                    }
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
function showAppConfirm(msg, onConfirm, customIcon = "❓") {
  const modal = document.createElement('div');
  modal.className = 'app-modal-overlay';
  modal.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); backdrop-filter:blur(5px); -webkit-backdrop-filter:blur(5px); display:flex; align-items:center; justify-content:center; z-index:10000; opacity:0; transition:opacity 0.3s;";
  
  modal.innerHTML = `
    <div style="background:#fff; width:90%; max-width:380px; border-radius:24px; padding:40px 30px; text-align:center; transform:scale(0.8); transition:transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1); box-shadow:0 20px 40px rgba(0,0,0,0.3); border-top: 6px solid #6b46c1;">
      <div style="font-size:4.5rem; margin-bottom:20px; animation: modalBounce 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);">${customIcon}</div>
      <h2 style="font-size:1.6rem; margin-bottom:12px; color:#1a202c; font-weight:800; font-family: sans-serif;">확인해주세요</h2>
      <p style="font-size:1.05rem; color:#4a5568; margin-bottom:30px; line-height:1.6; word-break:keep-all; font-family: sans-serif;">${msg}</p>
      <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px;">
        <button onclick="closeAppModal(this)" style="padding:18px; background:#f8fafc; color:#4a5568; border:1px solid #e2e8f0; border-radius:16px; font-size:1.1rem; font-weight:700; cursor:pointer; transition:all 0.2s;" onmouseover="this.style.background='#edf2f7'" onmouseout="this.style.background='#f8fafc'">취소</button>
        <button id="appConfirmOk" style="padding:18px; background:#6b46c1; color:#fff; border:none; border-radius:16px; font-size:1.1rem; font-weight:700; cursor:pointer; box-shadow: 0 8px 20px rgba(107,70,193,0.2); transition:all 0.2s;" onmouseover="this.style.filter='brightness(1.1)'" onmouseout="this.style.filter='none'">진행하기</button>
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
  // [v51.0] 출석체크, 회원가입, 회원재등록, 어드민 등 업무용/키오스크 페이지는 날씨 마법사 및 BGM을 원천 차단하여 성능을 보장합니다!
  if (
    path.includes('attendance') || 
    path.includes('registration') || 
    path.includes('renewal') || 
    path.includes('admin')
  ) {
    console.log("🚫 [Shared Environment] Weather and BGM disabled for utility/kiosk page.");
    const oldWrap = document.getElementById('village-weather-wrapper');
    if (oldWrap) oldWrap.remove();
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
  
  const isWindy = windSpeed >= 5.0;
  if (weather === 'sun' && !isWindy) return;
  
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
    `;
    document.head.appendChild(style);
  }
};

window.handleSharedBgm = function(enabled, url) {
  if (!url) return;
  
  let audio = document.getElementById('village-bgm-audio');
  if (!audio) {
    audio = document.createElement('audio');
    audio.id = 'village-bgm-audio';
    audio.loop = true;
    audio.src = url;
    document.body.appendChild(audio);
  } else if (audio.src !== url) {
    audio.src = url;
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
        audio.play().then(function() {
          btn.querySelector('i').style.color = '#e74c3c';
          btn.style.transform = 'scale(1.1) rotate(15deg)';
          localStorage.setItem('v44_bgm_user_play', 'true');
        }).catch(function(e) { console.log("BGM Play error:", e); });
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
      audio.play().then(function() {
        btn.querySelector('i').style.color = '#e74c3c';
        btn.style.transform = 'scale(1.1) rotate(15deg)';
        document.removeEventListener('click', startPlay);
        document.removeEventListener('touchstart', startPlay);
      }).catch(function(e) { console.log("BGM play failed on interaction:", e); });
    };
    
    audio.play().then(function() {
      btn.querySelector('i').style.color = '#e74c3c';
      btn.style.transform = 'scale(1.1) rotate(15deg)';
    }).catch(function(e) {
      console.log("Auto-play blocked by browser; registering interaction fallback.");
      document.addEventListener('click', startPlay);
      document.addEventListener('touchstart', startPlay);
    });
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
    history.back();
    console.log("[Global Backbutton] Modal deregistered. Stack size:", window.globalModalHistoryStack.length);
  }
};

window.addEventListener('popstate', function(e) {
  if (window.globalModalHistoryStack.length > 0) {
    // 뒤로가기 방어: 스택에서 가장 최근 등록된 닫기 콜백을 실행하여 모달을 닫아줍니다!
    const closeFn = window.globalModalHistoryStack.pop();
    if (closeFn) {
      closeFn(true); // popstate를 통한 물리적 닫힘으로 호출
    }
  } else {
    const path = window.location.pathname;
    const pageName = path.substring(path.lastIndexOf('/') + 1) || 'index.html';

    // 💡 [특수 설계] 예약, 출석, 가입, 관리자 화면 등 자체 히스토리 흐름을 관리하는 페이지는
    // 로컬 history state가 남아있으면 글로벌 퇴장 방지 결계의 가동을 일시 유예합니다!
    if (pageName === 'reservation.html' || pageName === 'attendance.html' || pageName === 'registration.html' || pageName === 'admin.html') {
      // 로컬 popstate 리스너가 흐름을 자연스럽게 되돌리도록 통제권을 양보함
      return;
    }

    // 열려있는 모달이 하나도 없을 때, 현재 화면에 알맞게 안전 퇴장 컨펌 작동!
    history.pushState(null, null, location.href); // 히스토리 밀림 복구
    
    if (pageName === 'oasis.html') {
      showAppConfirm("🌴 오아시스에서 퇴장하여 메인 광장으로 이동하시겠습니까?", function() {
        window.location.replace('index.html');
      }, "🌴");
    } else if (pageName === 'miracle.html') {
      showAppConfirm("⚔️ 모험가의 기록소에서 퇴장하여 메인 광장으로 이동하시겠습니까?", function() {
        window.location.replace('index.html');
      }, "⚔️");
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
