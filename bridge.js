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
      
      // [v44.220] 경로가 루트(/)인 경우를 대비해 절대 경로에 가깝게 리다이렉트
      const loginUrl = window.location.origin + window.location.pathname.replace(/\/[^\/]*$/, '/') + 'login.html';
      console.log("🛡️ Auth: Target redirect URL:", loginUrl);
      window.location.replace(loginUrl);
    }
  },
  logout: function() {
    if (confirm("🕌 지니 월드에서 퇴장하시겠습니까?")) {
      localStorage.removeItem('v44_user_name');
      localStorage.removeItem('v44_user_phone');
      localStorage.removeItem('v44_user_tier');
      window.location.replace('login.html');
    }
  }
};

// [v44.207] 모든 페이지 자동 체크 중단 -> 필요한 페이지(index, miracle 등)에서만 명시적으로 호출하도록 변경
// Auth.check(); 

/**
 * [공용] 환경에 맞는 페이지 이동 도우미
 */
function navigateTo(page, params = {}, openInNewTab = false) {
  const query = new URLSearchParams(params).toString();
  const isVercel = window.location.hostname.includes('vercel.app') || window.location.hostname.includes('localhost') || window.location.hostname.includes('127.0.0.1');
  
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
  modal.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); backdrop-filter:blur(5px); -webkit-backdrop-filter:blur(5px); display:flex; align-items:center; justify-content:center; z-index:10000; opacity:0; transition:opacity 0.3s;";
  
  const icon = type === "success" ? "✅" : (type === "error" ? "⚠️" : "ℹ️");
  const title = customTitle || (type === "success" ? "확인 완료" : (type === "error" ? "알림" : "안내"));
  const color = type === "success" ? "#38a169" : (type === "error" ? "#e53e3e" : "#6b46c1");

  const isLargeScreen = window.innerWidth > 700;
  // 단위를 rem 대신 px로 고정하여 브라우저 설정에 상관없이 크게 보이도록 함
  const modalWidth = isLargeScreen ? "800px" : "90%";
  const iconSize = isLargeScreen ? "120px" : "80px";
  const titleSize = isLargeScreen ? "45px" : "32px";
  const textSize = isLargeScreen ? "40px" : "24px";
  const btnSize = isLargeScreen ? "36px" : "24px";
  const btnPadding = isLargeScreen ? "35px" : "20px";

  modal.innerHTML = `
    <div style="background:#fff; width:95%; max-width:${modalWidth}; border-radius:40px; padding:${isLargeScreen ? '80px 50px' : '50px 30px'}; text-align:center; transform:scale(0.8); transition:transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1); box-shadow:0 40px 80px rgba(0,0,0,0.5); border-top: 15px solid ${color};">
      <div style="font-size:${iconSize}; margin-bottom:30px; animation: modalBounce 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);">${icon}</div>
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
function showAppConfirm(msg, onConfirm) {
  const modal = document.createElement('div');
  modal.className = 'app-modal-overlay';
  modal.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); backdrop-filter:blur(5px); -webkit-backdrop-filter:blur(5px); display:flex; align-items:center; justify-content:center; z-index:10000; opacity:0; transition:opacity 0.3s;";
  
  modal.innerHTML = `
    <div style="background:#fff; width:90%; max-width:380px; border-radius:24px; padding:40px 30px; text-align:center; transform:scale(0.8); transition:transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1); box-shadow:0 20px 40px rgba(0,0,0,0.3); border-top: 6px solid #6b46c1;">
      <div style="font-size:4.5rem; margin-bottom:20px; animation: modalBounce 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);">❓</div>
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
 * [v44.210] 앱 비정상 종료 방지 (휴대폰 뒤로가기 트랩)
 */
window.addEventListener('load', function() {
  if (window.history && window.history.pushState) {
    // 페이지 진입 시 가짜 히스토리를 하나 추가합니다.
    window.history.pushState('app-back-trap', null, null);
    
    window.addEventListener('popstate', function(e) {
      // 사용자가 휴대폰 뒤로가기 버튼을 누르면 이 이벤트가 발생합니다.
      // 앱이 꺼지는 것을 막기 위해 즉시 가짜 히스토리를 다시 추가합니다.
      window.history.pushState('app-back-trap', null, null);
      
      // 사용자에게 종료(또는 이전 화면 이동) 의사를 묻습니다.
      showAppConfirm("앱을 종료(또는 이전 화면으로 이동)하시겠습니까?", function() {
        // 사용자가 '진행하기'를 누르면,
        // 강제로 히스토리를 2칸 뒤로 되돌려 실제 뒤로가기(또는 앱 종료)를 수행합니다.
        window.history.go(-2);
      });
    });
  }
});

/**
 * [v44.218] 글로벌 로딩 시스템
 */
function showLoading(msg) {
  const overlay = document.getElementById('v-loading');
  const msgEl = document.getElementById('v-loading-msg');
  if (overlay && msgEl) {
    msgEl.innerText = msg || "기록을 동기화 중...";
    overlay.style.display = 'flex';
  }
}
function hideLoading() {
  const overlay = document.getElementById('v-loading');
  if (overlay) overlay.style.display = 'none';
}
