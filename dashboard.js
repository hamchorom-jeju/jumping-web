/**
 * Nohyung Village Dashboard Logic (v44.110 - Premium Ranking Ticker)
 * Features: Badge-based Hall of Fame, Smooth Vertical Animation, v44.0 Immutable Base
 */

const Village = {
    perspective: 'weekly',
    user: {
        name: "회원",
        tier: "꿈나무 요정",
        totalScore: 84200,
        rank: 12,
        stats: {
            weekly: { health: 450, perf: 680, def: 320 },
            monthly: { health: 1200, perf: 2400, def: 950 }
        },
        max: {
            weekly: { health: 600, perf: 1000, def: 700 },
            monthly: { health: 2400, perf: 4000, def: 2800 }
        },
        tiers: [
            { name: "씨앗 🌱", min: 0 }, { name: "새싹 🌿", min: 1001 }, { name: "나무 🌳", min: 3001 },
            { name: "꽃 🌸", min: 8001 }, { name: "꿈나무 요정 🧚‍♂️", min: 15001 },
            { name: "전설의 점퍼 👑", min: 30001 }, { name: "지니 월드 수호신 🌌", min: 60001 }
        ],
        water: 1.2,
        habits: [
            { id: 'h1', title: '모닝 티', done: false, base: 2 },
            { id: 'h2', title: '모닝 스트레칭', done: false, base: 2 },
            { id: 'h3', title: '베지 퍼스트', done: false, base: 2 },
            { id: 'h4', title: '슬로우 치잉', done: false, base: 2 },
            { id: 'h5', title: '7,000보 달성', done: false, base: 3 },
            { id: 'h6', title: '스테어 마법', done: false, base: 2 },
            { id: 'h7', title: '나이트 컷', done: false, base: 10 },
            { id: 'h8', title: '굿 슬립', done: false, base: 2 },
            { id: 'h9', title: '셀프 칭찬', done: false, base: 2 },
            { id: 'plus', title: '✨ 미라클 플러스', done: false, base: 0 }
        ]
    },

    habitData: {
        h1: { title: "모닝 티", icon: "🍵", guide: "기상 직후 따뜻한 물이나 차 한 잔은 밤새 잠들었던 장기를 깨우고 신진대사의 점화를 유도합니다. 체온보다 약간 높은 온도의 미온수를 권장하며, 이는 대사 촉진 및 독소 배출의 소중한 첫 단계가 됩니다.", link: "miracle.html?cat=habit&item=h1" },
        h2: { title: "모닝 스트레칭", icon: "🧘", guide: "기상 후 5분 내외의 가벼운 전신 스트레칭은 신체의 가동 범위를 확보하고 근육에 산소를 공급합니다. 목, 어깨, 허리 위주의 부드러운 이완으로 활기찬 하루를 시작해 보세요.", link: "miracle.html?cat=habit&item=h2" },
        h3: { title: "베지 퍼스트", icon: "🥗", guide: "식사 시 채소를 먼저 섭취하면 인슐린 스파이크(급격한 혈당 상승)를 효과적으로 방지할 수 있습니다. 첫 젓가락은 무조건 채소류부터! 이후 단백질과 탄수화물 순으로 섭취하여 지방 축적을 최소화하세요.", link: "miracle.html?cat=habit&item=h3" },
        h4: { title: "슬로우 치잉", icon: "🦷", guide: "한 입당 20회 이상 천천히 씹는 습관은 소화를 돕고 뇌가 포만감을 느낄 시간을 충분히 부여합니다. 식사 시간을 최소 20분 이상 유지하는 것은 과식 방지의 핵심입니다.", link: "#", single: true },
        h5: { title: "7,000보 달성", icon: "👟", guide: "일상 활동량을 정량적으로 확보하여 기초 대사량을 유지하고 혈당 조절에 기여합니다. 스마트폰 기준으로 7,000보를 달성하며 꾸준한 에너지 소모를 실천해 보세요.", link: "miracle.html?cat=habit&item=h5" },
        h6: { title: "스테어 마법", icon: "🪜", guide: "3층 이하의 낮은 층수는 계단을 이용함으로써 하체 근력을 강화하고 유산소 효과를 얻습니다. 특히 기록소(아카이브)에 사진 인증을 남기시면 강력한 보너스 점수가 부여됩니다.", link: "miracle.html?cat=habit&item=h6" },
        h7: { title: "나이트 컷", icon: "🌙", guide: "밤 20시 이후의 금식은 신체가 소화가 아닌 '지방 연소와 세포 재생'에 집중하게 만듭니다. 지금 이 시간부터 아무것도 먹지 않겠다고 [오아시스]에 다짐의 선언을 남겨보세요!\n\n\"저 지금부터는 아무것도 안 먹어요.. 약속합니다!!\" 라는 한마디가 강력한 수호의 시작입니다.", link: "oasis.html" },
        h8: { title: "굿 슬립", icon: "💤", guide: "세포가 재생되고 성장 호르몬이 활발히 분비되는 자정(24:00) 전 취침으로 신체 회복을 최적화하세요. 충분한 수면은 식욕 억제 호르몬인 렙틴의 분비를 도와 다이어트를 수월하게 만듭니다.", link: "#", single: true },
        h9: { title: "셀프 칭찬", icon: "👏", guide: "오늘 하루도 노력한 나자신을 위해 따뜻한 한마디를 해주며 셀프 허그를 해주세요.\n\"오늘도 수고했어 영희야!\"\n오아시스 게시판에 셀프칭찬글도 남겨보세요.", link: "oasis.html" },
        plus: { title: "미라클 플러스", icon: "✨", guide: "건강 습관 외에 인생을 풍요롭게 만드는 사소하지만 위대한 승리들을 인증해주세요.\n새벽기상, 독서인증, 환경수호 활동(플로깅 등) 등...\n\n📸 아카이브에 등록시 5점이 지급됩니다.\n(일 최대 1건 등록 가능)", link: "miracle.html?cat=plus", single: true }
    },

    // 🏆 Structured Ranking Data (v44.110)
    rankings: [
        { type: "금주 랭킹", badge: "weekly", content: "<span>👑 체력왕: 홍길동</span><span>🏅 미션왕: 김개똥</span><span>🛡️ 수호왕: 이성실</span>" },
        { type: "월간 랭킹", badge: "monthly", content: "<span>🔥 다이어트킹: 박지니</span><span>🚀 미션 선두: 최열정</span>" },
        { type: "토탈 랭킹", badge: "total", content: "<span>🏆 1위: 전설모험가</span><span>🥈 2위: 꾸준지존</span><span>🥉 3위: 열정맨</span>" }
    ],
    currentRankIndex: 0,

    quests: {
        sync: { title: "클럽 동기화", icon: "⚡", guide: "클럽 출석 점수를 반영하시겠어요?\n오늘 클럽에 출석 하셨다면 15포인트,\n운동량에 따라 최대 20포인트가 반영됩니다.", btn: "동기화 수행", link: "#", single: true },
        visit: { title: "방문 인증", icon: "📸", guide: "오늘 클럽에 출석하셨나요?\n그렇다면 방문 인증을 남겨보세요\n아카이브에 인증시 15점이 추가됩니다.", btn: "인증하러 가기", link: "miracle.html?cat=visit", single: false },
        meal: { title: "식단 인증", icon: "🍱", guide: "꾸준한 식단 기록은 강력한 변화의 열쇠입니다.\n아카이브에 인증시 최대 30점이 지급됩니다.", btn: "인증하러 가기", link: "miracle.html?cat=meal", single: false },
        water: { title: "워터 헌터", icon: "💧", guide: "수분 섭취량 만큼 게이지를 조정해보세요.\n섭취량에 따라 점수를 차등지급합니다. (최대 20점)\n아카이브에 인증시 15점이 추가됩니다.", btn: "인증하러 가기", link: "miracle.html?cat=water", single: false },
        bonus: { title: "보너스 퀘스트", icon: "✨", guide: "돌발 미션을 수행하시겠어요?\n아카이브로 이동해 인증을 남기실 수 있습니다.\n아카이브 인증시 15점이 추가됩니다.", btn: "인증하러 가기", link: "miracle.html?cat=bonus", single: false }
    },

    showLoading(msg) {
        const overlay = document.getElementById('v-loading');
        const msgEl = document.getElementById('v-loading-msg');
        if (overlay && msgEl) {
            // [v44.229] If already showing, just update message to prevent flicker
            msgEl.innerText = msg || "기록을 동기화 중...";
            overlay.style.display = 'flex';
        }
    },
    hideLoading() {
        const overlay = document.getElementById('v-loading');
        if (overlay) overlay.style.display = 'none';
    },

    /**
     * [v44.229] 프리미엄 뒤로가기 핸들러
     */
    lastBackTime: 0,
    handleBackButton() {
        // 1. 열려있는 모달이 있는지 확인 (Vanilla JS 및 프리미엄 모달 공통)
        const premiumModal = document.querySelector('.app-modal-overlay');
        const habitModal = document.getElementById('habit-modal');

        if (premiumModal) {
            // 프리미엄 알림/확인창 닫기
            const btn = premiumModal.querySelector('button');
            if (btn) btn.click();
            history.pushState(null, null, location.href); 
            return;
        }

        if (habitModal && habitModal.style.display === 'flex') {
            // 습관 가이드 모달 닫기
            this.closeModal();
            history.pushState(null, null, location.href); 
            return;
        }

        // 2. 모달이 없으면 종료 안내 토스트
        const now = Date.now();
        if (now - this.lastBackTime < 2000) {
            alert("지니 월드(노형 빌리지)를 이용해주셔서 감사합니다.\n로그아웃 버튼으로 안전하게 퇴장하실 수 있습니다. ✨");
        } else {
            this.lastBackTime = now;
            this.showToast("한 번 더 누르면 종료 안내가 표시됩니다.");
            history.pushState(null, null, location.href); 
        }
    },

    showToast(msg) {
        const toast = document.createElement('div');
        toast.innerText = msg;
        toast.style = "position:fixed; bottom:100px; left:50%; transform:translateX(-50%); background:rgba(0,0,0,0.7); color:#fff; padding:10px 20px; border-radius:20px; z-index:9999; font-size:0.8rem; font-weight:800;";
        document.body.appendChild(toast);
        setTimeout(() => toast.remove(), 2000);
    },

    init() {
        console.log("v44.229 Real Data Sync Initialized.");
        
        // [v44.229] 뒤로가기 방지용 히스토리 추가
        history.pushState(null, null, location.href);
        window.onpopstate = () => this.handleBackButton();

        this.loadRealData();
        this.renderAll();
        this.updateEvolution();
        this.startTicker();
        this.bindEvents();
    },

    loadRealData() {
        const params = new URLSearchParams(window.location.search);
        let phone = (params.get('phone') || '').trim();
        
        // [v44.169] URL에 정보가 없으면 브라우저 저장소(localStorage)에서 복원 시도
        if (!phone) {
            phone = (localStorage.getItem('v44_user_phone') || '').trim();
        }
        
        if (!phone) return;

        // [v44.193] 로딩 시작
        this.showLoading("📜 마을 기록을 불러오고 있습니다...");

        if (typeof google !== 'undefined' && google.script && google.script.run) {
            google.script.run
                .withSuccessHandler(res => {
                    this.hideLoading(); // 로딩 종료
                    if (res && res.success) {
                        this.user.name = res.name;
                        this.user.tier = res.tier;
                        localStorage.setItem('v44_user_tier', res.tier);
                        this.user.nextTier = res.nextTier;
                        this.user.evolution = res.evolution;
                        this.user.totalScore = res.totalScore;
                        this.user.rank = res.rank;

                        // [v45.0] 오늘 완료한 항목 체크박스 복구
                        if (res.doneList && res.doneList.length > 0) {
                            console.log("[v45.0] Restoring doneList:", res.doneList);
                            this.user.habits.forEach(h => {
                                // 예: doneList에 "모닝 티 완료"가 있으면 h.title "모닝 티"와 매칭
                                if (res.doneList.some(item => item.indexOf(h.title) > -1)) {
                                    h.done = true;
                                }
                            });
                        }

                        if (res.stats) {
                            this.user.stats = res.stats;
                            if (res.stats.targets) this.user.max = res.stats.targets;
                        }
                        this.renderAll();
                        this.updateEvolution();
                        if (res.isFirstLoginToday) this.showLoginReward();
                    }
                })
                .withFailureHandler(() => this.hideLoading())
                .getUserDashboardData({ phone: phone });
        }
    },

    // [v44.170] 프리미엄 로그인 보상 알림창
    showLoginReward() {
        // [v44.172] 담백하고 정중한 환영 메시지로 수정
        const data = {
            icon: "✨",
            title: "오늘의 첫 접속 보너스",
            guide: "오늘도 노형 빌리지를 찾아주셔서 감사합니다!\n첫 접속 보상으로 **+5 EXP**가 적립되었습니다.\n\n오늘도 건강하고 활기찬 하루 되세요!",
            btn: "반가워요! 🙌"
        };
        
        document.getElementById('modal-habit-icon').innerText = data.icon;
        document.getElementById('modal-habit-title').innerText = data.title;
        document.getElementById('modal-habit-guide').innerHTML = data.guide.replace(/\*\*(.*?)\*\*/g, '<b style="color:var(--gold);">$1</b>');
        
        const cancelBtn = document.getElementById('modal-cancel-btn');
        const confirmBtn = document.getElementById('modal-confirm-btn');
        
        cancelBtn.style.display = 'none';
        confirmBtn.innerText = data.btn;
        confirmBtn.style.flex = "1";
        
        confirmBtn.onclick = () => this.closeModal();
        document.getElementById('habit-modal').style.display = 'flex';
    },

    openModal(key, type = 'quest') {
        const data = (type === 'quest') ? this.quests[key] : this.habitData[key];
        const habit = (type === 'habit') ? this.user.habits.find(h => h.id === key) : null;
        if (!data) return;
        
        document.getElementById('modal-habit-icon').innerText = data.icon;
        document.getElementById('modal-habit-title').innerText = data.title || data.meaning;
        
        let guideText = data.guide;
        if (type === 'habit') {
            if (key === 'plus') {
            } else if (data.single) {
                guideText += `\n\n🌿 수호 완료 시 ${habit.base}점 지급`;
            } else if (key === 'h7') {
                guideText += `\n\n🌿 수호 완료 시 ${habit.base}점 지급\n🌵 다짐의 선언 게시글 등록 시 5점 추가`;
            } else if (key === 'h9') {
                guideText += `\n\n🌿 수호 완료 시 ${habit.base}점 지급\n👏 셀프칭찬 등록 시 5점 추가`;
            } else {
                guideText += `\n\n🌿 수호 완료 시 ${habit.base}점 지급\n📸 아카이브 인증 시 5점 추가`;
            }
        }
        document.getElementById('modal-habit-guide').innerText = guideText;
        const cancelBtn = document.getElementById('modal-cancel-btn');
        const confirmBtn = document.getElementById('modal-confirm-btn');
        
        if ((type === 'quest' && data.single) || (type === 'habit' && data.single)) {
            cancelBtn.style.display = 'none';
            confirmBtn.innerText = (key === 'plus') ? "인증하러 가기" : (type === 'habit' ? "수호 완료" : data.btn);
            confirmBtn.style.flex = "1";
        } else {
            cancelBtn.style.display = 'block';
            cancelBtn.innerText = (type === 'habit') ? "체크만 하기" : "나중에";
            confirmBtn.innerText = (type === 'habit') ? ((key === 'h7' || key === 'h9') ? "게시판 이동" : "인증하러 가기") : data.btn;
            confirmBtn.style.flex = "1.5";
        }
        
        cancelBtn.onclick = () => {
            // [v44.229] h7(나이트컷), h9(셀프칭찬)도 '체크만 하기' 가능하도록 수정
            if (type === 'habit' && habit) Village.applyHabitCheck(key, false);
            Village.closeModal();
        };

        confirmBtn.onclick = () => {
            if (data.link !== "#") {
                if (type === 'habit' && habit && key !== 'plus') Village.applyHabitCheck(key, false);
                const finalLink = (type === 'habit' && key !== 'h7' && key !== 'h9' && key !== 'plus') 
                    ? `miracle.html?tab=habit&item=${key}` : data.link.replace('cat=', 'tab=');
                location.href = finalLink;
            } else if (key === 'sync') {
                Village.syncClubRecordActual();
            } else if (type === 'habit' && data.single) {
                Village.applyHabitCheck(key, false);
            }
            Village.closeModal();
        };

        document.getElementById('habit-modal').style.display = 'flex';
    },

    applyHabitCheck(id, withAuth) {
        const habit = this.user.habits.find(h => h.id === id);
        if (habit && !habit.done) {
            // [v44.229] 낙관적 업데이트: 로딩바 없이 즉시 화면 업데이트
            habit.done = true;
            const points = habit.base + (withAuth ? 5 : 0);
            
            // 로컬 수치 즉시 반영 (심리적 만족감)
            this.user.totalScore += points;
            this.user.stats.weekly.def += points;
            this.renderAll();
            this.updateEvolution();

            if (typeof google !== 'undefined' && google.script && google.script.run) {
                const params = new URLSearchParams(window.location.search);
                let phone = (params.get('phone') || '').trim();
                if (!phone) phone = (localStorage.getItem('v44_user_phone') || '').trim();

                // 서버에는 조용히 기록 (로딩바 없음)
                // [v45.0] 인증 여부에 따른 action 추가 (완료 vs 인증)
                google.script.run
                    .withSuccessHandler(() => {
                        console.log(`[v45.0] Habit ${id} silently recorded.`);
                    })
                    .recordActivityLog({
                        phone: phone, 
                        name: this.user.name, 
                        type: "습관", 
                        item: habit.title, 
                        score: points,
                        action: withAuth ? "인증" : "완료",
                        statType: "def"
                    });
            }
        }
    },

    /**
     * [v44.229] 전체 데이터 동기화 (마스터 플랜 지침: 수동 갱신)
     */
    syncData() {
        this.loadRealData();
    },

    closeModal() { document.getElementById('habit-modal').style.display = 'none'; },
    openQuestModal(key) { this.openModal(key, 'quest'); },

    syncClubRecordActual() {
        const params = new URLSearchParams(window.location.search);
        let phone = (params.get('phone') || '').trim();
        if (!phone) phone = (localStorage.getItem('v44_user_phone') || '').trim();

        if (typeof google !== 'undefined' && google.script && google.script.run) {
            this.showLoading("🏡 클럽 기록을 가져오고 있습니다...");
            google.script.run
                .withSuccessHandler(res => {
                    this.hideLoading();
                    if (res && res.success) {
                        alert(`🏡 [클럽 동기화 완료!]\n총 ${res.points} EXP가 반영되었습니다.\n(방문보너스: 15 + 운동타임: ${res.timePoints})`);
                        this.loadRealData();
                    } else {
                        alert(`❌ 동기화 실패: ${res.error}`);
                    }
                })
                .withFailureHandler(() => this.hideLoading())
                .syncClubRecord({ phone: phone });
        }
    },

    startTicker() {
        const container = document.getElementById('ranking-ticker');
        if (!container) return;

        const renderItem = (idx) => {
            const r = this.rankings[idx];
            return `
                <div class="v-ranking-item" style="transform: translateY(100%); opacity: 0;">
                    <div class="v-ranking-badge ${r.badge}">${r.type}</div>
                    <div class="v-ranking-content">${r.content}</div>
                </div>
            `;
        };

        container.innerHTML = this.rankings.map((_, i) => renderItem(i)).join('');
        const items = container.querySelectorAll('.v-ranking-item');
        
        const showItem = (idx) => {
            items.forEach((item, i) => {
                if (i === idx) {
                    item.style.transform = 'translateY(0)';
                    item.style.opacity = '1';
                } else {
                    item.style.transform = 'translateY(-100%)';
                    item.style.opacity = '0';
                }
            });
        };

        showItem(0);
        setInterval(() => {
            this.currentRankIndex = (this.currentRankIndex + 1) % this.rankings.length;
            showItem(this.currentRankIndex);
        }, 5000);
    },

    renderAll() {
        document.getElementById('user-name').innerText = this.user.name;
        document.getElementById('user-tier').innerText = this.user.tier; // 칭호 업데이트
        document.getElementById('total-score').innerText = (this.user.totalScore || 0).toLocaleString();
        document.getElementById('current-rank').innerText = this.user.rank;
        document.getElementById('water-val').innerText = `${this.user.water}L / 2.0L`;

        // [v44.228] 3D 성장배지 스프라이트 업데이트
        const badgeEl = document.getElementById('user-badge-sprite');
        if (badgeEl && this.user.tier) {
            badgeEl.className = 'v-badge-sprite'; // 초기화
            
            const tier = this.user.tier;
            if (tier.includes("씨앗")) badgeEl.classList.add('v-badge-seed');
            else if (tier.includes("새싹")) badgeEl.classList.add('v-badge-sprout');
            else if (tier.includes("나무")) badgeEl.classList.add('v-badge-tree');
            else if (tier.includes("꽃")) badgeEl.classList.add('v-badge-flower');
            else if (tier.includes("요정")) badgeEl.classList.add('v-badge-fairy');
            else if (tier.includes("점퍼")) badgeEl.classList.add('v-badge-legend');
            else if (tier.includes("수호신")) badgeEl.classList.add('v-badge-guardian');
        }
        
        const view = this.perspective;
        const currentData = this.user.stats[view];
        const maxData = this.user.max[view];
        
        // [v44.167] 항목별 개별 목표치 적용
        this.updateGauge('health', currentData.health, maxData.health);
        this.updateGauge('perf', currentData.perf, maxData.perf);
        this.updateGauge('def', currentData.def, maxData.def);
        this.renderHabits();
    },

    updateGauge(id, val, max) {
        if (!max) max = 1000; // Defensive fallback
        if (typeof max === 'object') max = max[id] || 1000;
        
        // [v44.167] "현재 / 목표" 형식의 정밀 수치 표기
        const typeNames = { health: '❤️ 체력', perf: '🗡️ 수행력', def: '🛡️ 방어력' };
        document.getElementById(`${id}-val`).innerText = `${Math.floor(val).toLocaleString()} / ${max.toLocaleString()}`;
        
        const percent = Math.min((val / max) * 100, 100);
        document.getElementById(`${id}-bar`).style.width = `${percent}%`;
        document.getElementById(`${id}-label`).innerText = `${this.perspective === 'weekly' ? '주간' : '월간'} ${typeNames[id]} 달성률: ${Math.round(percent)}%`;
    },

    updateEvolution() {
        // [v44.167] 서버에서 계산된 v39 공식 진화도 적용
        const evoBar = document.getElementById('evo-bar');
        const evoLabel = document.getElementById('evo-percent');
        const nextLabel = document.getElementById('next-tier-label');
        
        if (evoBar) evoBar.style.width = `${this.user.evolution || 0}%`;
        if (evoLabel) evoLabel.innerText = `${this.user.evolution || 0}%`;
        if (nextLabel) nextLabel.innerText = `Next: ${this.user.nextTier || '새싹 🌿'}`;
    },

    setPerspective(view) {
        this.perspective = view;
        const btnW = document.getElementById('btn-weekly');
        const btnM = document.getElementById('btn-monthly');
        if (view === 'weekly') {
            btnW.style.background = '#fff'; btnW.style.boxShadow = '0 2px 5px rgba(0,0,0,0.1)';
            btnM.style.background = 'transparent'; btnM.style.boxShadow = 'none';
        } else {
            btnM.style.background = '#fff'; btnM.style.boxShadow = '0 2px 5px rgba(0,0,0,0.1)';
            btnW.style.background = 'transparent'; btnW.style.boxShadow = 'none';
        }
        this.renderAll();
    },

    renderHabits() {
        const container = document.getElementById('habit-list-container');
        if (!container) return;
        container.innerHTML = this.user.habits.map(h => `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:12px; background:${h.id === 'plus' ? '#fff9c4' : '#fff'}; border-radius:15px; margin-bottom:8px; border:1px solid var(--v-border);">
                <span style="font-size:0.9rem; font-weight:800; color:var(--v-wood);">${h.title}</span>
                <div onclick="Village.openModal('${h.id}', 'habit')" style="width:28px; height:28px; border:2px solid var(--v-border); border-radius:8px; cursor:pointer; display:flex; align-items:center; justify-content:center; background:${h.done ? 'var(--def)' : 'transparent'};">
                    ${h.done ? '✅' : ''}
                </div>
            </div>
        `).join('');
    },

    bindEvents() {
        const slider = document.getElementById('water-range');
        if (slider) {
            slider.addEventListener('input', (e) => {
                const val = (e.target.value / 10).toFixed(1);
                document.getElementById('water-val').innerText = `${val}L / 2.0L`;
            });
        }
    },

    triggerSuddenMission() { document.getElementById('sudden-mission-bar').style.display = 'block'; }
};

window.onload = () => Village.init();
window.syncClubRecord = () => Village.openQuestModal('sync');
window.triggerSuddenMission = () => Village.openQuestModal('bonus');
