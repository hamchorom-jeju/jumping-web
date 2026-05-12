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
function showAppAlert(msg, type = "success") {
  const modal = document.createElement('div');
  modal.className = 'app-modal-overlay';
  modal.style.cssText = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.7); backdrop-filter:blur(5px); -webkit-backdrop-filter:blur(5px); display:flex; align-items:center; justify-content:center; z-index:10000; opacity:0; transition:opacity 0.3s;";
  
  const icon = type === "success" ? "✅" : (type === "error" ? "⚠️" : "ℹ️");
  const title = type === "success" ? "처리 완료" : (type === "error" ? "알림" : "안내");
  const color = type === "success" ? "#38a169" : (type === "error" ? "#e53e3e" : "#6b46c1");

  const isLargeScreen = window.innerWidth > 700;
  const modalWidth = isLargeScreen ? "650px" : "420px";
  const iconSize = isLargeScreen ? "8rem" : "5rem";
  const titleSize = isLargeScreen ? "2.8rem" : "1.8rem";
  const textSize = isLargeScreen ? "2rem" : "1.4rem";
  const btnSize = isLargeScreen ? "2rem" : "1.3rem";
  const btnPadding = isLargeScreen ? "28px" : "22px";

  modal.innerHTML = `
    <div style="background:#fff; width:92%; max-width:${modalWidth}; border-radius:32px; padding:${isLargeScreen ? '70px 45px' : '50px 35px'}; text-align:center; transform:scale(0.8); transition:transform 0.3s cubic-bezier(0.34, 1.56, 0.64, 1); box-shadow:0 35px 70px rgba(0,0,0,0.4); border-top: 12px solid ${color};">
      <div style="font-size:${iconSize}; margin-bottom:25px; animation: modalBounce 0.6s cubic-bezier(0.34, 1.56, 0.64, 1);">${icon}</div>
      <h2 style="font-size:${titleSize}; margin-bottom:15px; color:#1a202c; font-weight:800; font-family: sans-serif;">${title}</h2>
      <p style="font-size:${textSize}; color:#2d3748; margin-bottom:40px; line-height:1.6; word-break:keep-all; font-family: sans-serif; font-weight:500;">${msg}</p>
      <button onclick="closeAppModal(this)" 
        style="width:100%; padding:${btnPadding}; background:${color}; color:#fff; border:none; border-radius:24px; font-size:${btnSize}; font-weight:800; cursor:pointer; box-shadow: 0 10px 30px rgba(0,0,0,0.2);">
        확인
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
