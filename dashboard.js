/**
 * Nohyung Village Dashboard Logic (v44.110 - Premium Ranking Ticker)
 * Features: Badge-based Hall of Fame, Smooth Vertical Animation, v44.0 Immutable Base
 */

const Village = {
    perspective: 'weekly',
    user: {
        name: "모험가",
        tier: "씨앗 🌱",
        totalScore: 0,
        rank: "-",
        stats: {
            weekly: { health: 0, perf: 0, def: 0 },
            monthly: { health: 0, perf: 0, def: 0 }
        },
        max: {
            weekly: { health: 1500, perf: 1000, def: 500 },
            monthly: { health: 6000, perf: 4000, def: 2000 }
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
        h2: { title: "모닝 스트레칭", icon: "🧘", guide: "기상 후 5분 내외의 가벼운 전신 스트레칭은 신체의 가동 범위를 확보하고 근육에 산소를 공급합니다. 목, 어깨, 허리 위주의 부드러운 이완으로 활기찬 하루를 시작해 보세요.", link: "#", single: true },
        h3: { title: "베지 퍼스트", icon: "🥗", guide: "식사 시 채소를 먼저 섭취하면 인슐린 스파이크(급격한 혈당 상승)를 효과적으로 방지할 수 있습니다. 첫 젓가락은 무조건 채소류부터! 이후 단백질과 탄수화물 순으로 섭취하여 지방 축적을 최소화하세요.", link: "#", single: true },
        h4: { title: "슬로우 치잉", icon: "🦷", guide: "한 입당 20회 이상 천천히 씹는 습관은 소화를 돕고 뇌가 포만감을 느낄 시간을 충분히 부여합니다. 식사 시간을 최소 20분 이상 유지하는 것은 과식 방지의 핵심입니다.", link: "#", single: true },
        h5: { title: "7,000보 달성", icon: "👟", guide: "일상 활동량을 정량적으로 확보하여 기초 대사량을 유지하고 혈당 조절에 기여합니다. 스마트폰 기준으로 7,000보를 달성하며 꾸준한 에너지 소모를 실천해 보세요.", link: "miracle.html?cat=habit&item=h5" },
        h6: { title: "스테어 마법", icon: "🪜", guide: "3층 이하의 낮은 층수는 계단을 이용함으로써 하체 근력을 강화하고 유산소 효과를 얻습니다. 특히 기록소(아카이브)에 사진 인증을 남기시면 강력한 보너스 점수가 부여됩니다.", link: "miracle.html?cat=habit&item=h6" },
        h7: { title: "나이트 컷", icon: "🌙", guide: "밤 20시 이후의 금식은 신체가 소화가 아닌 '지방 연소와 세포 재생'에 집중하게 만듭니다. 오늘 저녁 나이트 컷을 약속하신다면 아래에서 [체크만 하기]를 누르시거나, [저녁식단 인증하러 가기]를 통해 따뜻한 차나 물 사진을 인증하여 자동으로 나이트 컷을 완료하세요! 🍵", link: "miracle.html?tab=meal" },
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
        sync: { title: "클럽 동기화", icon: "⚡", guide: "클럽 출석 점수를 반영하시겠어요?\n오늘 클럽에 출석 하셨다면 15포인트,\n운동량에 따라 최대 20포인트가 반영됩니다.", btn: "동기화 진행", link: "#", single: true },
        visit: { title: "방문 인증", icon: "📸", guide: "오늘 클럽에 출석하셨나요?\n그렇다면 방문 인증을 남겨보세요\n아카이브에 인증시 15점이 추가됩니다.", btn: "인증하러 가기", link: "miracle.html?cat=visit", single: false },
        meal: { title: "식단 인증", icon: "🍱", guide: "꾸준한 식단 기록은 강력한 변화의 열쇠입니다.\n아카이브에 인증시 최대 30점이 지급됩니다.", btn: "인증하러 가기", link: "miracle.html?cat=meal", single: false },
        water: { title: "워터 헌터", icon: "💧", guide: "수분 섭취량 만큼 게이지를 조정해보세요.\n섭취량에 따라 점수를 차등지급합니다. (최대 20점)\n아카이브에 인증시 15점이 추가됩니다.", btn: "인증하러 가기", link: "miracle.html?cat=water", single: false },
        bonus: { title: "보너스 퀘스트", icon: "✨", guide: "돌발 미션을 실천하시겠어요?\n아카이브로 이동해 인증을 남기실 수 있습니다.\n아카이브 인증시 15점이 추가됩니다.", btn: "인증하러 가기", link: "miracle.html?cat=bonus", single: false }
    },

    showLoading(msg) {
        console.log("⏳ [Silent Sync] " + (msg || "기록을 동기화 중..."));
        // [v50.9] 원장님 지침에 따라 대시보드 조작을 방해하는 어두운 장막(로딩 마스크)을 완벽하게 영구 무력화합니다!
        const overlay = document.getElementById('v-loading');
        if (overlay) overlay.style.display = 'none';
    },
    hideLoading() {
        const overlay = document.getElementById('v-loading');
        if (overlay) overlay.style.display = 'none';
    },
    showSyncToast() {
        const id = 'sync-toast';
        if (document.getElementById(id)) return;
        const toast = document.createElement('div');
        toast.id = id;
        toast.innerHTML = '🔄 최신 데이터로 동기화 중...';
        toast.style = "position:fixed; top:14px; right:14px; background:rgba(30,30,40,0.82); color:#fff; padding:8px 16px; border-radius:20px; z-index:9999; font-size:0.75rem; font-weight:700; backdrop-filter:blur(6px); transition:opacity 0.4s;";
        document.body.appendChild(toast);
    },
    hideSyncToast() {
        const toast = document.getElementById('sync-toast');
        if (toast) {
            toast.style.opacity = '0';
            setTimeout(() => toast.remove(), 400);
        }
    },

    init() {
        console.log("v44.229 Real Data Sync Initialized.");

        // [perf] 낙관적 렌더링: localStorage 캐시가 있으면 즉시 화면에 먼저 표시
        const cached = localStorage.getItem('v44_dashboard_cache');
        let cacheLoaded = false;
        if (cached) {
            try {
                const c = JSON.parse(cached);
                this.user.name = c.name || this.user.name;
                this.user.tier = c.tier || this.user.tier;
                this.user.totalScore = c.totalScore || this.user.totalScore;
                this.user.rank = c.rank || this.user.rank;
                this.user.stats = c.stats || this.user.stats;
                this.user.max = c.max || this.user.max;
                if (c.habits) {
                    c.habits.forEach(h => {
                        const found = this.user.habits.find(x => x.id === h.id);
                        if (found) found.done = h.done;
                    });
                }
                cacheLoaded = true;
            } catch(e) {
                localStorage.removeItem('v44_dashboard_cache'); // 손상된 캐시 제거
            }
        }
        this.renderAll();
        this.updateEvolution();
        this.loadRealData(cacheLoaded);
        this.loadRankingTickerData();
        this.startTicker();
        this.bindEvents();
    },

    loadRealData(cacheLoaded = false) {
        const params = new URLSearchParams(window.location.search);
        let phone = (params.get('phone') || '').trim();
        
        // [v44.169] URL에 정보가 없으면 브라우저 저장소(localStorage)에서 복원 시도
        if (!phone) {
            phone = (localStorage.getItem('v44_user_phone') || '').trim();
        }
        
        if (!phone) return;

        if (typeof google !== 'undefined' && google.script && google.script.run) {
            // 캐시가 로드된 경우: 작은 토스트로 동기화 안내, 캐시 없으면: 전체 로딩 오버레이 표시
            if (cacheLoaded) {
                this.hideLoading(); // ⚡ 이미 캐시가 로드되었으므로 풀스크린 로딩을 즉시 해제합니다!
                this.showSyncToast();
            } else {
                this.showLoading("📜 마을 기록을 불러오고 있습니다...");
            }

            google.script.run
                .withSuccessHandler(res => {
                    this.hideLoading();
                    this.hideSyncToast(); // 동기화 토스트 제거
                    if (res && res.success) {
                        this.user.name = res.name;
                        this.user.tier = res.tier;
                        localStorage.setItem('v44_user_tier', res.tier);
                        this.user.nextTier = res.nextTier;
                        this.user.evolution = res.evolution;
                        this.user.totalScore = res.totalScore;
                        this.user.rank = res.rank;

                        // [v45.0] 오늘 완료한 항목 체크박스 복구
                        // [v47.0] 날이 바뀐 경우 이전 상태를 리셋하기 위해 모든 습관의 done 상태를 초기화합니다.
                        this.user.habits.forEach(h => {
                            h.done = false;
                        });
                        this.user.doneList = res.doneList || [];
                        if (res.doneList && res.doneList.length > 0) {
                            console.log("[v45.0] Restoring doneList:", res.doneList);
                            this.user.habits.forEach(h => {
                                if (res.doneList.some(item => item.indexOf(h.title) > -1)) {
                                    h.done = true;
                                }
                            });
                        }

                        if (res.stats) {
                            this.user.stats = res.stats;
                        }
                        // [v45.9] 주간/월간 목표치 각각 저장
                        this.user.max = {
                            weekly: res.weeklyTargets || { health: 1500, perf: 1000, def: 500 },
                            monthly: res.monthlyTargets || { health: 6000, perf: 4000, def: 2000 }
                        };

                        // [perf] 서버 응답을 localStorage에 캐시 저장 (다음 방문 시 즉시 렌더링용)
                        try {
                            localStorage.setItem('v44_dashboard_cache', JSON.stringify({
                                name: res.name,
                                tier: res.tier,
                                totalScore: res.totalScore,
                                rank: res.rank,
                                stats: res.stats,
                                max: this.user.max,
                                habits: this.user.habits.map(h => ({ id: h.id, done: h.done }))
                            }));
                        } catch(e) { /* 스토리지 가득 찬 경우 무시 */ }

                        this.renderAll();
                        this.updateEvolution();
                        
                        // [v59.0] 글로벌 우편함 데이터 실시간 동기화
                        if (window.syncMailboxWithDashboardData) {
                            window.syncMailboxWithDashboardData(res);
                        }

                        if (res.quests) {
                            try {
                                this.renderQuestWidgets(res.quests);
                            } catch (e) {
                                console.error("Error rendering quest widgets:", e);
                            }
                        }
                        if (res.inactiveDays !== undefined) {
                            this.renderInactivityDebuff(res.inactiveDays, res.inactivityPenalty);
                        }
                        if (res.isFirstLoginToday) this.showLoginReward();
                        
                        // [v55.0] 실시간 릴레이 칭찬 수신 여부 및 칭찬 바톤 주자 팝업 가동!
                        if (res.praiseBatonSender) {
                            this.showPraiseBatonNotice(res.praiseBatonSender);
                        } else if (res.praiseNotice) {
                            this.showPraiseWizardNotice(res.praiseNotice);
                        }
                        
                        // 🌦️ 마을 기후 및 배경음악 환경 실시간 동기화 반영
                        if (res.villageSettings) {
                            this.applyVillageEnvironment(res.villageSettings);
                        }
                                                // 📢 전령의 기둥 공지 단독 바인딩 (돌발 퀘스트 롤링 제외) [v66.0]
                        if (res.pillarNotice) {
                            const noticeEl = document.getElementById('village-notice-banner');
                            if (noticeEl) {
                                let notices = [];
                                if (Array.isArray(res.pillarNotice)) {
                                    notices = [...res.pillarNotice];
                                } else if (res.pillarNotice && typeof res.pillarNotice === 'object') {
                                    notices = [res.pillarNotice];
                                }

                                if (notices.length === 0) {
                                    noticeEl.innerHTML = `<div class="v-banner-notice-inner"><span class="v-notice-badge">📢 마을공지</span><span class="v-notice-text">오늘도 건강한 하루 되세요!</span></div><span class="v-banner-badge">전체보기 <i class="fa-solid fa-chevron-right"></i></span>`;
                                    noticeEl.dataset.targetUrl = 'notice.html';
                                } else if (notices.length === 1) {
                                    const item = notices[0];
                                    const title = item.title || item.content || '';
                                    noticeEl.innerHTML = `<div class="v-banner-notice-inner"><span class="v-notice-badge">📢 마을공지</span><span class="v-notice-text">${title.trim()}</span></div><span class="v-banner-badge">전체보기 <i class="fa-solid fa-chevron-right"></i></span>`;
                                    noticeEl.dataset.targetUrl = 'notice.html';
                                } else {
                                    let curIdx = 0;
                                    
                                    const renderInitial = () => {
                                        const item = notices[curIdx];
                                        const title = item.title || item.content || '';
                                        
                                        noticeEl.innerHTML = `<div class="v-banner-notice-inner"><span class="v-notice-badge">📢 마을공지</span><span class="v-notice-text">${title.trim()}</span></div><span class="v-banner-badge">전체보기 <i class="fa-solid fa-chevron-right"></i></span>`;
                                        noticeEl.dataset.targetUrl = 'notice.html';
                                        
                                        if (window.noticeInterval) clearTimeout(window.noticeInterval);
                                        window.noticeInterval = setTimeout(transitionToNext, 5000);
                                    };
                                    
                                    const transitionToNext = () => {
                                        const nextIdx = (curIdx + 1) % notices.length;
                                        const nextItem = notices[nextIdx];
                                        const nextTitle = nextItem.title || nextItem.content || '';
                                        
                                        const textEl = noticeEl.querySelector('.v-notice-text');
                                        if (textEl) {
                                            textEl.style.transition = 'opacity 0.25s ease';
                                            textEl.style.opacity = '0';
                                            
                                            setTimeout(() => {
                                                textEl.textContent = nextTitle.trim();
                                                textEl.style.opacity = '1';
                                                
                                                curIdx = nextIdx;
                                                
                                                if (window.noticeInterval) clearTimeout(window.noticeInterval);
                                                window.noticeInterval = setTimeout(transitionToNext, 5000);
                                            }, 250);
                                        } else {
                                            curIdx = nextIdx;
                                            renderInitial();
                                        }
                                    };
                                    
                                    renderInitial();
                                }
                                
                                noticeEl.onclick = () => {
                                    const url = noticeEl.dataset.targetUrl || 'notice.html';
                                    location.href = url;
                                };
                            }
                        }
                    }
                })
                .withFailureHandler(() => { this.hideLoading(); this.hideSyncToast(); })
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
        
        // [v47.0] 이미 오늘 완료된 습관인 경우 추가 획득/진입을 차단하고 친절하게 안내합니다.
        if (type === 'habit' && habit && habit.done) {
            this.showToast(`✨ 이미 오늘 ${habit.title} 수호가 완료되었습니다!`);
            return;
        }
        
        document.getElementById('modal-habit-icon').innerText = data.icon;
        document.getElementById('modal-habit-title').innerText = data.title || data.meaning;
        
        let guideText = data.guide;
        if (type === 'habit') {
            if (key === 'plus') {
            } else if (data.single) {
                guideText += `\n\n🌿 수호 완료 시 ${habit.base}점 지급`;
            } else if (key === 'h7') {
                guideText += `\n\n🌿 수호 완료 시 ${habit.base}점 지급\n🍵 저녁 단식(물/차) AI 인증 시 10점 자동 연계 지급`;
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
            confirmBtn.innerText = (type === 'habit') ? (key === 'h7' ? "저녁식단 인증하러 가기" : key === 'h9' ? "게시판 이동" : "인증하러 가기") : data.btn;
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
        if (window.registerModalOpen) {
            window.registerModalOpen(Village.closeModal);
        }
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

    closeModal(isPopstate) {
        const modal = document.getElementById('habit-modal');
        if (modal) modal.style.display = 'none';
        if (!isPopstate && window.registerModalClose) {
            window.registerModalClose(Village.closeModal);
        }
    },
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
                        showAppAlert(`🏡 [클럽 동기화 완료!]\n총 ${res.points} EXP가 반영되었습니다.\n(방문보너스: 20 + 운동타임: ${res.timePoints})`, "success");
                        this.loadRealData();
                    } else {
                        showAppAlert(`❌ 동기화 실패: ${res.error}`, "error");
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
        if (window.rankingTickerInterval) clearInterval(window.rankingTickerInterval);
        window.rankingTickerInterval = setInterval(() => {
            this.currentRankIndex = (this.currentRankIndex + 1) % this.rankings.length;
            showItem(this.currentRankIndex);
        }, 5000);
    },

    /**
     * [v60.0] 실시간/시상식 랭킹 데이터를 티커에 정밀 반영 및 동적 월명 렌더링
     */
    updateRankingsFromData(data) {
        if (!data) return;

        const currentMonth = new Date().getMonth() + 1;
        const formattedRankings = [];
        
        const today = new Date();
        const dayOfWeek = today.getDay(); // 4 = Thursday
        const dateOfMonth = today.getDate(); // 1, 2, 3

        // 큰 EXP 수치를 k단위로 축약하여 프리미엄 룩 완성 (예: 38,400 -> 38.4k)
        const formatScore = (score, isCount = false, unit = "p") => {
            if (isCount) return `${score}${unit}`;
            if (score >= 10000) {
                return `${(score / 1000).toFixed(1)}k${unit}`;
            }
            return `${score.toLocaleString()}${unit}`;
        };

        // 상위 3명 순위를 메달 뱃지와 함께 템플릿 스트링으로 가공 (멀티라인 2줄 전격 교체)
        const getTop3Html = (list, isCount = false, unit = "p") => {
            if (!list || list.length === 0) return "<div class='v-ranking-line line-1'>데이터 집계 중...</div>";
            
            const items = list.slice(0, 3).map((item, idx) => {
                const scoreVal = item.score !== undefined ? item.score : item.count;
                if (idx === 0) {
                    return `<span class="v-ranker-span">🥇 1.${item.name}(${formatScore(scoreVal, isCount, unit)})<span class="v-live-badge">실시간<span class="v-pulse-dot"></span></span></span>`;
                } else {
                    return `<span class="v-ranker-span">${idx + 1}.${item.name}(${formatScore(scoreVal, isCount, unit)})</span>`;
                }
            });
            
            const line1 = `<div class="v-ranking-line line-1">${items[0] || ''}</div>`;
            const line2 = items.length > 1 
                ? `<div class="v-ranking-line line-2">${items[1]}${items[2] ? items[2] : ''}</div>`
                : '';
                
            return `<div class="v-ranking-content-multiline">${line1}${line2}</div>`;
        };

        // 확정된 아카이브 랭킹 데이터를 시상식 뷰로 가공하는 헬퍼 함수
        const getConfirmedTop3Html = (winnersStr) => {
            if (!winnersStr) return "<div class='v-ranking-line line-1'>시상 기록 데이터가 없습니다.</div>";
            
            // 🥇 왕관 이모지 제거 후 세로 바 기준으로 split
            const cleanStr = winnersStr.replace("🥇 ", "");
            const tokens = cleanStr.split("|").map(t => t.trim());
            
            const w1 = tokens[0] || "-";
            const w2 = tokens[1] || "-";
            const w3 = tokens[2] || "-";
            
            const line1 = `<div class="v-ranking-line line-1"><span class="v-ranker-span">🥇 ${w1}<span class="v-live-badge award" style="background:#4f46e5; box-shadow: 0 0 8px rgba(79, 70, 229, 0.4);">🏆 시상<span class="v-pulse-dot" style="background:#fff;"></span></span></span></div>`;
            const line2 = `<div class="v-ranking-line line-2"><span class="v-ranker-span">${w2}</span><span class="v-ranker-span">${w3}</span></div>`;
            
            return `<div class="v-ranking-content-multiline">${line1}${line2}</div>`;
        };

        // (1) 주간 랭킹 TOP 3 (목요일: 지난주 확정 시상 배너로 자동 스위칭)
        const isWeeklyConfirmedDay = (dayOfWeek === 4); // 목요일
        const hasWeeklyArchive = (data.archive && data.archive.weekly && data.archive.weekly.length > 0);
        
        if (isWeeklyConfirmedDay && hasWeeklyArchive) {
            const latestWeeklyArchive = data.archive.weekly[0];
            formattedRankings.push({
                type: `🏆 ${latestWeeklyArchive.period} 시상`,
                badge: "weekly",
                content: getConfirmedTop3Html(latestWeeklyArchive.winners)
            });
        } else if (data.weekly && data.weekly.length > 0) {
            formattedRankings.push({
                type: "🏆 주간랭킹",
                badge: "weekly",
                content: getTop3Html(data.weekly, false, "p")
            });
        }

        // (2) 월간 랭킹 TOP 3 (1~3일: 지난달 확정 시상 배너로 자동 스위칭)
        const isMonthlyConfirmedDay = (dateOfMonth >= 1 && dateOfMonth <= 3); // 매월 1일 ~ 3일
        const hasMonthlyArchive = (data.archive && data.archive.monthly && data.archive.monthly.length > 0);
        
        if (isMonthlyConfirmedDay && hasMonthlyArchive) {
            const latestMonthlyArchive = data.archive.monthly[0];
            formattedRankings.push({
                type: `📅 ${latestMonthlyArchive.period} 시상`,
                badge: "monthly",
                content: getConfirmedTop3Html(latestMonthlyArchive.winners)
            });
        } else if (data.monthly && data.monthly.length > 0) {
            formattedRankings.push({
                type: "📅 월간랭킹",
                badge: "monthly",
                content: getTop3Html(data.monthly, false, "p")
            });
        }

        // (3) 토탈 랭킹 TOP 3 (변함없는 누적 랭킹)
        if (data.total && data.total.length > 0) {
            formattedRankings.push({
                type: "💫 토탈랭킹",
                badge: "total",
                content: getTop3Html(data.total, false, "p")
            });
        }

        // (4) 당월 출석왕 실시간 라이브 닷 레이스 (1~3일: 지난달 최종 확정 출석왕 시상 배너로 자동 스위칭)
        const hasAttendanceArchive = (data.archive && data.archive.attendance && data.archive.attendance.length > 0);
        
        if (isMonthlyConfirmedDay && hasAttendanceArchive) {
            const latestAttArchive = data.archive.attendance[0];
            formattedRankings.push({
                type: `🏃 ${latestAttArchive.period} 출석왕`,
                badge: "weekly", // 녹색 배지
                content: getConfirmedTop3Html(latestAttArchive.winners)
            });
        } else if (data.mvp && data.mvp.monthlyAttendance && data.mvp.monthlyAttendance.length > 0) {
            formattedRankings.push({
                type: "🏃 출석왕",
                badge: "weekly", // 활기찬 초록색 라이브 배지 연출
                content: getTop3Html(data.mvp.monthlyAttendance, true, "회")
            });
        }

        if (formattedRankings.length > 0) {
            this.rankings = formattedRankings;
            this.refreshTicker();
        }
    },

    refreshTicker() {
        this.currentRankIndex = 0;
        this.startTicker();
    },

    /**
     * [v5.0] 실시간 랭킹 티커 데이터 호출 (구글 서버 API 및 로컬 모의 백엔드 하이브리드)
     */
    loadRankingTickerData() {
        if (typeof google !== 'undefined' && google.script && google.script.run) {
            google.script.run
                .withSuccessHandler(res => {
                    if (res && res.success && res.data) {
                        this.updateRankingsFromData(res.data);
                    }
                })
                .getHallOfFameData();
        } else {
            console.warn("Running in local browser environment. Injecting premium ranking mock data...");
            setTimeout(() => {
                const MOCK_RANKING_DATA = {
                    weekly: [
                        { name: "김영희", score: 1450 },
                        { name: "이철수", score: 1280 },
                        { name: "박민수", score: 1100 }
                    ],
                    monthly: [
                        { name: "이철수", score: 5800 },
                        { name: "김영희", score: 5200 },
                        { name: "최미진", score: 4800 }
                    ],
                    total: [
                        { name: "박민수", score: 38400 },
                        { name: "이철수", score: 36200 },
                        { name: "김영희", score: 32400 }
                    ],
                    mvp: {
                        monthlyAttendance: [
                            { name: "김영희", count: 24 },
                            { name: "이철수", count: 22 },
                            { name: "박민수", count: 21 }
                        ]
                    },
                    archive: {
                        weekly: [
                            { period: "5월 4주", winners: "🥇 1.김영희(1,450p) | 2.이철수(1,280p) | 3.박민수(1,100p)" }
                        ],
                        monthly: [
                            { period: "5월", winners: "🥇 1.이철수(5,800p) | 2.김영희(5,200p) | 3.최미진(4,800p)" }
                        ],
                        attendance: [
                            { period: "5월", winners: "🥇 1위 김영희(24회) | 🥈 2위 이철수(22회) | 🥉 3위 박민수(21회)" }
                        ]
                    }
                };
                this.updateRankingsFromData(MOCK_RANKING_DATA);
            }, 1000);
        }
    },

    renderAll() {
        const nameEl = document.getElementById('user-name');
        if (nameEl) nameEl.innerText = this.user.name;

        const tierEl = document.getElementById('user-tier');
        if (tierEl) tierEl.innerText = this.user.tier;

        const scoreEl = document.getElementById('total-score');
        if (scoreEl) scoreEl.innerText = (this.user.totalScore || 0).toLocaleString();

        // [v68.0] 주간/월간 관점에 따른 총점 계산 및 각 토글 버튼 텍스트 동적 업데이트 (에러 완벽 방지, 대칭 정밀 유지)
        const toggleView = this.perspective || 'weekly';
        const toggleStats = this.user && this.user.stats ? this.user.stats : null;
        const toggleCurrentData = toggleStats && toggleStats[toggleView] ? toggleStats[toggleView] : { health: 0, perf: 0, def: 0 };
        const totalPerspectiveScore = (toggleCurrentData.health || 0) + (toggleCurrentData.perf || 0) + (toggleCurrentData.def || 0);

        const btnW = document.getElementById('btn-weekly');
        const btnM = document.getElementById('btn-monthly');
        
        if (btnW && btnM) {
            if (toggleView === 'weekly') {
                btnW.innerText = `주간 누적 ${Math.floor(totalPerspectiveScore).toLocaleString()}점`;
                btnM.innerText = `월간`;
            } else {
                btnW.innerText = `주간`;
                btnM.innerText = `월간 누적 ${Math.floor(totalPerspectiveScore).toLocaleString()}점`;
            }
        }

        const rankEl = document.getElementById('current-rank');
        if (rankEl) rankEl.innerText = this.user.rank;

        const waterEl = document.getElementById('water-val');
        if (waterEl) waterEl.innerText = `${this.user.water}L / 2.0L`;

        // [v44.228] 3D 성장배지 스프라이트 업데이트
        const badgeEl = document.getElementById('user-badge-sprite');
        if (badgeEl && this.user.tier) {
            badgeEl.className = 'v-badge-sprite'; // 초기화
            
            const tier = this.user.tier;
            if (tier.includes("수호신")) badgeEl.classList.add('v-badge-guardian');
            else if (tier.includes("점퍼")) badgeEl.classList.add('v-badge-legend');
            else if (tier.includes("요정")) badgeEl.classList.add('v-badge-fairy');
            else if (tier.includes("꽃")) badgeEl.classList.add('v-badge-flower');
            else if (tier.includes("나무")) badgeEl.classList.add('v-badge-tree');
            else if (tier.includes("새싹")) badgeEl.classList.add('v-badge-sprout');
            else if (tier.includes("씨앗")) badgeEl.classList.add('v-badge-seed');
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
        const typeNames = { health: '❤️ 체력', perf: '🗡️ 실천력', def: '🔮 회복력' };
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

    triggerSuddenMission() { Village.openQuestModal('bonus'); },

    renderQuestWidgets(quests) {
        if (!quests) return;
        
        // 💡 [iOS Safari 날짜 파싱 호환성 헬퍼]
        // 아이폰 Safari는 YYYY-MM-DD 형태의 대시(-) 날짜 문자열을 new Date로 파싱할 때 
        // Invalid Date 에러(NaN)를 내며 자바스크립트 전체를 중단시키는 결함이 있어,
        // 대시를 슬래시(/)로 치환하여 파싱 호환성을 보장합니다.
        const safeParseDate = (dateStr) => {
            if (!dateStr) return new Date();
            const cleanStr = String(dateStr).replace(/-/g, '/');
            return new Date(cleanStr);
        };
        
        // ⚡ [v66.0] 오늘의 돌발 퀘스트 전용 프리미엄 아코디언 배너 동적 데이터 바인딩
                const suddenBanner = document.getElementById('sudden-quest-accordion-banner');
        if (suddenBanner) {
            if (quests.todayQuest) {
                const q = quests.todayQuest;
                const isCompleted = (this.user && this.user.doneList && Array.isArray(this.user.doneList) && q && q.title) ?
                    this.user.doneList.some(item => typeof item === 'string' && (item.indexOf(q.title) > -1 || item.indexOf('돌발퀘스트') > -1)) : false;
                
                if (isCompleted) {
                    // [UX 개선] 오늘의 돌발 미션을 완료한 경우, 대시보드 공간 절약을 위해 배너를 완전히 숨깁니다!
                    suddenBanner.style.display = 'none';
                } else {
                    suddenBanner.style.display = 'flex';
                    
                    const badgeEl = document.getElementById('sudden-badge-type');
                    const titleEl = document.getElementById('sudden-title-text');
                    const descEl = document.getElementById('sudden-desc-text');
                    const rewardEl = document.getElementById('sudden-reward-text');
                    const methodEl = document.getElementById('sudden-method-text');
                    const actionBtn = document.getElementById('sudden-action-btn');
                    
                    // 테마 초기화
                    suddenBanner.classList.remove('theme-rose', 'theme-purple', 'theme-green');
                    
                    const methodStr = String(q.method || '사진').trim();
                    if (methodStr === '게시판') {
                        suddenBanner.classList.add('theme-purple');
                        if (badgeEl) badgeEl.innerText = '📝 게시판돌발';
                        if (actionBtn) {
                            actionBtn.innerHTML = '<i class="fa-solid fa-pen-to-square"></i> 게시판 작성하러 가기';
                            actionBtn.dataset.targetUrl = `oasis.html?write=sudden-quest&qtitle=${encodeURIComponent(q.title)}`;
                        }
                    } else {
                        suddenBanner.classList.add('theme-rose');
                        if (badgeEl) badgeEl.innerText = '📸 사진돌발';
                        if (actionBtn) {
                            actionBtn.innerHTML = '<i class="fa-solid fa-camera"></i> 사진 인증하러 가기';
                            actionBtn.dataset.targetUrl = 'miracle.html?tab=quest&cat=bonus';
                        }
                    }
                    if (titleEl) titleEl.innerText = `⚡ ${q.title}`;
                    if (descEl) descEl.innerText = q.description || '';
                    if (rewardEl) rewardEl.innerText = `+${q.score || 15} EXP`;
                    if (methodEl) methodEl.innerText = `${methodStr} 인증`;
                    if (actionBtn) {
                        actionBtn.disabled = false;
                        actionBtn.style.opacity = '1';
                    }
                }
            } else {
                suddenBanner.style.display = 'none';
            }
        }
        

        // 3. 글리코겐 클리어 방패 퀘스트 위젯
        const glyWidget = document.getElementById('glycogen-quest-widget');
        if (glyWidget) {
            if (quests.glycogenQuest) {
                glyWidget.style.display = 'flex';
                
                // 남은 시간 계산 (시간 단위) - safeParseDate 적용
                const deadlineTime = safeParseDate(quests.glycogenQuest.deadlineStr).getTime();
                const nowTime = Date.now();
                const diffMs = deadlineTime - nowTime;
                const hoursLeft = Math.max(0, Math.ceil(diffMs / (1000 * 60 * 60)));
                
                const deadlineEl = document.getElementById('glycogen-quest-deadline');
                if (deadlineEl) {
                    if (hoursLeft > 24) {
                        const daysLeft = Math.floor(hoursLeft / 24);
                        deadlineEl.innerText = `마감 D-${daysLeft}일 (${hoursLeft}시간 남음)`;
                    } else {
                        deadlineEl.innerText = `마감 임박! ${hoursLeft}시간 남음`;
                    }
                }
                
                // 출석 둥근 원형 표시
                const ratioEl = document.getElementById('glycogen-quest-ratio');
                if (ratioEl) ratioEl.innerText = `${quests.glycogenQuest.progress} / 3`;
                
                const circlesContainer = document.getElementById('glycogen-circles');
                if (circlesContainer) {
                    let circlesHtml = '';
                    const progress = quests.glycogenQuest.progress || 0;
                    for (let i = 0; i < 3; i++) {
                        if (i < progress) {
                            circlesHtml += `<div style="width: 16px; height: 16px; border-radius: 50%; background: #be123c; border: 2px solid #be123c;"></div>`;
                        } else {
                            circlesHtml += `<div style="width: 16px; height: 16px; border-radius: 50%; background: transparent; border: 2px solid #be123c;"></div>`;
                        }
                    }
                    circlesContainer.innerHTML = circlesHtml;
                }
            } else {
                glyWidget.style.display = 'none';
            }
        }
        
        // 4. 요요 방패 버프 위젯
        const shieldWidget = document.getElementById('yoyo-shield-buff-widget');
        if (shieldWidget) {
            if (quests.shield && quests.shield.active) {
                shieldWidget.style.display = 'flex';
                
                // safeParseDate 적용
                const expireTime = safeParseDate(quests.shield.expireStr).getTime();
                const nowTime = Date.now();
                const diffMs = expireTime - nowTime;
                const daysLeft = Math.max(0, Math.ceil(diffMs / (1000 * 60 * 60 * 24)));
                
                const timeEl = document.getElementById('yoyo-shield-buff-time');
                if (timeEl) {
                    timeEl.innerText = `남은 기한: ${daysLeft}일 (D-${daysLeft})`;
                }
            } else {
                shieldWidget.style.display = 'none';
            }
        }
    },

    renderInactivityDebuff(inactiveDays, inactivityPenalty) {
        const widget = document.getElementById('inactivity-debuff-widget');
        const daysEl = document.getElementById('debuff-days');
        const penaltyEl = document.getElementById('debuff-penalty');
        
        if (widget && daysEl && penaltyEl) {
            if (inactivityPenalty > 0) {
                daysEl.innerText = inactiveDays;
                penaltyEl.innerText = inactivityPenalty.toLocaleString();
                widget.style.display = 'flex';
            } else {
                widget.style.display = 'none';
            }
        }
    },

    showPraiseBatonNotice(senderName) {
        if (document.getElementById('v-praise-baton-modal')) return;
        const modal = document.createElement('div');
        modal.id = 'v-praise-baton-modal';
        modal.style = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(15, 118, 110, 0.45); backdrop-filter:blur(8px); display:flex; align-items:center; justify-content:center; z-index:99999;";
        modal.innerHTML = `
            <div style="background: linear-gradient(135deg, #ffffff 0%, #fffbeb 100%); width:90%; max-width:380px; padding:30px; border-radius:30px; border:2.5px solid #eab308; box-shadow:0 20px 50px rgba(0,0,0,0.15); text-align:center;">
                <div style="font-size: 3rem; margin-bottom: 15px; animation: bounce 2s infinite;">🏃‍♂️🔗</div>
                <h3 style="font-size: 1.15rem; font-weight: 900; color: #854d0e; margin: 0 0 10px 0;">칭찬 릴레이 바톤 터치!</h3>
                <p style="font-size: 0.82rem; font-weight: 800; color: #334155; line-height: 1.6; word-break: keep-all; margin-bottom: 22px;">
                    축하합니다! 🌟<br>
                    <strong>[${senderName}]</strong> 님이 당신을 지목하여 칭찬의 바톤을 넘겨주셨습니다!<br><br>
                    <span style="color: #d97706; font-weight: 900;">지금 바로 다음 주자를 지목해 끊어지지 않는 칭찬의 고리를 완성해 주세요! 🌴</span>
                </p>
                <div style="display:flex; gap:10px;">
                    <button onclick="document.getElementById('v-praise-baton-modal').remove()" style="flex:1; background:rgba(0,0,0,0.06); border:none; padding:12px 0; border-radius:15px; font-size:0.8rem; font-weight:900; color:#475569; cursor:pointer;">확인</button>
                    <button onclick="location.href='oasis.html'" style="flex:1.5; background:linear-gradient(135deg, #eab308 0%, #ca8a04 100%); border:none; padding:12px 0; border-radius:15px; font-size:0.8rem; font-weight:900; color:#fff; cursor:pointer; box-shadow:0 4px 12px rgba(234,179,8,0.25);">바톤 이어받기</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    },

    showPraiseWizardNotice(message) {
        if (document.getElementById('v-praise-magic-modal')) return;
        const modal = document.createElement('div');
        modal.id = 'v-praise-magic-modal';
        modal.style = "position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(15, 118, 110, 0.45); backdrop-filter:blur(8px); display:flex; align-items:center; justify-content:center; z-index:99999;";
        modal.innerHTML = `
            <div style="background: linear-gradient(135deg, #ffffff 0%, #f0fdf4 100%); width:90%; max-width:380px; padding:30px; border-radius:30px; border:2px solid #0d9488; box-shadow:0 20px 50px rgba(0,0,0,0.15); text-align:center;">
                <div style="font-size: 3rem; margin-bottom: 15px; animation: bounce 2s infinite;">💌</div>
                <h3 style="font-size: 1.15rem; font-weight: 900; color: #0f766e; margin: 0 0 10px 0;">칭찬의 마법 배달!</h3>
                <p style="font-size: 0.82rem; font-weight: 800; color: #334155; line-height: 1.6; word-break: keep-all; margin-bottom: 22px;">
                    ${message}<br><br>
                    <span style="color: #0d9488; font-weight: 900;">지금 바로 오아시스 탭으로 이동하여 따뜻한 칭찬의 글을 확인해 보세요! 🌴</span>
                </p>
                <div style="display:flex; gap:10px;">
                    <button onclick="document.getElementById('v-praise-magic-modal').remove()" style="flex:1; background:rgba(0,0,0,0.06); border:none; padding:12px 0; border-radius:15px; font-size:0.8rem; font-weight:900; color:#475569; cursor:pointer;">확인</button>
                    <button onclick="location.href='oasis.html'" style="flex:1.5; background:linear-gradient(135deg, #0d9488 0%, #0f766e 100%); border:none; padding:12px 0; border-radius:15px; font-size:0.8rem; font-weight:900; color:#fff; cursor:pointer; box-shadow:0 4px 12px rgba(15,118,110,0.25);">오아시스로 이동</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    },

    applyVillageEnvironment(settings) {
        if (!settings) return;
        this.lastSettings = settings; // 토글 시 재참조용으로 보관
        
        // [v65.0] 마스터 환경 마법 일괄 제어 토글 처리 (원장님 일괄 차단 지시 수용)
        const magicEnabled = settings.magicEnabled === undefined || settings.magicEnabled === null || settings.magicEnabled.toString().toLowerCase() !== 'false';
        if (!magicEnabled) {
            console.log("🚫 [Dashboard Environment] Globally disabled by the Chief.");
            const oldWrap = document.getElementById('village-weather-wrapper');
            if (oldWrap) oldWrap.remove();
            
            // BGM 일시정지 및 버튼 숨김
            const bgmBtn = document.getElementById('village-bgm-toggle');
            if (bgmBtn) bgmBtn.style.display = 'none';
            const audio = document.getElementById('village-bgm-audio');
            if (audio) audio.pause();
            
            // 제주 싱크 날씨 배너 가리기
            const jejuBar = document.getElementById('jeju-sync-weather-bar');
            if (jejuBar) jejuBar.style.display = 'none';
            
            // 날씨 토글 스위치 비주얼 비활성화
            const weatherBtn = document.getElementById('jeju-weather-toggle-btn');
            if (weatherBtn) {
                weatherBtn.innerHTML = '🚫 비활성화됨';
                weatherBtn.classList.add('disabled');
            }
            return;
        }
        
        // [v48.0] 회원별 날씨 효과 온/오프 상태 로드 및 스위치 UI 동기화
        const weatherDisabled = localStorage.getItem('village_weather_disabled') === 'true';
        const btn = document.getElementById('jeju-weather-toggle-btn');
        if (btn) {
            if (weatherDisabled) {
                btn.innerHTML = '🚫 켜기';
                btn.classList.add('disabled');
            } else {
                btn.innerHTML = '👁️ 끄기';
                btn.classList.remove('disabled');
            }
        }
        
        // 온/오프 상태에 따라 파티클 렌더링 값 분기
        const weather = weatherDisabled ? 'sun' : (settings.resolvedWeather || settings.weather || 'sun');
        const windSpeed = weatherDisabled ? 0 : (parseFloat(settings.realJejuWind) || 0);
        this.renderWeatherParticles(weather, windSpeed);
        
        const bgmEnabled = settings.bgmEnabled && settings.bgmEnabled.toString().toLowerCase() === 'true';
        const bgmUrl = settings.bgmUrl || '';
        this.handleBgm(bgmEnabled, bgmUrl);

        // [v48.0] 실시간 제주시 노형동 기상 싱크 배너 노출 (수동 날씨 모드일 때도 온/오프 버튼 제어를 위해 노출!)
        const jejuBar = document.getElementById('jeju-sync-weather-bar');
        const jejuText = document.getElementById('jeju-sync-weather-text');
        if (jejuBar && jejuText) {
            const origWeather = settings.weather || 'sun';
            
            // [v51.0] 기상 상태와 무관하게 날씨 배너는 항상 우아하게 노출하여 자리를 채웁니다.
            jejuBar.style.display = 'flex';
                
            let weatherEmoji = '☀️';
            let weatherName = '맑음';
            
            const displayWeather = settings.resolvedWeather || origWeather;
            
            if (displayWeather === 'rain') { weatherEmoji = '🌧️'; weatherName = '비'; }
            else if (displayWeather === 'snow') { weatherEmoji = '❄️'; weatherName = '눈'; }
            else if (displayWeather === 'blossom') { weatherEmoji = '🌸'; weatherName = '벚꽃 흩날림'; }
            else if (displayWeather === 'leaves') { weatherEmoji = '🍁'; weatherName = '낙엽 낙하'; }
            else if (displayWeather === 'sun') { weatherEmoji = '☀️'; weatherName = '맑음'; }
            
            if (settings.realJejuTemp !== undefined) {
                const temp = settings.realJejuTemp;
                const origWind = parseFloat(settings.realJejuWind) || 0;
                
                let windStatus = '';
                if (origWind >= 14.0) {
                    windStatus = ` ⚠️ <strong>태풍급 강풍 주의!</strong>`;
                    jejuBar.classList.add('wind-gale');
                } else if (origWind >= 5.0) {
                    windStatus = ` 🍃 <strong>강한 제주의 바람 부는 중!</strong>`;
                    jejuBar.classList.add('wind-gale');
                } else {
                    windStatus = ` 🍃 산들바람`;
                    jejuBar.classList.remove('wind-gale');
                }
                
                jejuText.innerHTML = `현재 제주시 노형동은 <strong>${weatherEmoji} ${weatherName} (${temp}°C)</strong>, 풍속 <strong>${origWind} m/s</strong>${windStatus}`;
            } else {
                // 수동 세팅 모드인 경우
                jejuBar.classList.remove('wind-gale');
                jejuText.innerHTML = `마을 기후 마법 작동 중: <strong>${weatherEmoji} ${weatherName}</strong> (수동 설정)`;
            }
        }
    },
    
    renderWeatherParticles(weather, windSpeed = 0) {
        const oldWrap = document.getElementById('village-weather-wrapper');
        if (oldWrap) oldWrap.remove();
        
        // 바람 세기 정의: 5.0 m/s 이상이면 강풍(windy) 비주얼 모드 발동!
        const isWindy = windSpeed >= 5.0;
        
        const wrapper = document.createElement('div');
        wrapper.id = 'village-weather-wrapper';
        wrapper.style = "position:fixed; top:0; left:0; width:100%; height:100%; pointer-events:none; z-index:9999; overflow:hidden;";
        document.body.appendChild(wrapper);
        
        // 1. 강풍 상태(Jeju Windy)일 때 은은하게 흐르는 '바람 줄기(Wind Wisps)' 이펙트 추가!
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
                
                // 풍속에 비례해서 더 빠르게 휭휭 날아감!
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
        
        // 2. 날씨 효과 렌더링 (비, 눈, 벚꽃, 낙엽)
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
                
                // 강풍이 불면 떨어지는 속도를 2배 빠르게 조절!
                let duration = Math.random() * 5 + 5;
                if (isWindy) duration = duration * 0.45;
                
                const delay = Math.random() * 8;
                
                if (weather === 'rain') {
                    p.style.width = '1.5px';
                    p.style.height = (Math.random() * 20 + 15) + 'px';
                    p.style.background = 'rgba(174, 219, 255, 0.6)';
                    
                    // 강풍일 때는 빗줄기가 휘몰아치듯 더 심하게 꺾임! (35도 vs 15도)
                    const angle = isWindy ? 35 : 15;
                    p.style.transform = `rotate(${angle}deg)`;
                } else {
                    p.innerText = chars[weather] || '';
                }
                
                // 강풍일 때는 더 멀리 옆으로 날아가도록 전용 windy 애니메이션 적용!
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
                @keyframes blow-wind {
                    0% { transform: translateX(-250px); }
                    100% { transform: translateX(110vw); }
                }
                @keyframes fall-rain {
                    to { transform: translateY(105vh) rotate(15deg); }
                }
                @keyframes fall-rain-windy {
                    to { transform: translateY(105vh) translateX(35vw) rotate(35deg); }
                }
                @keyframes fall-snow {
                    0% { transform: translateY(-20px) translateX(0) rotate(0deg); }
                    50% { transform: translateY(50vh) translateX(20px) rotate(180deg); }
                    100% { transform: translateY(105vh) translateX(-10px) rotate(360deg); }
                }
                @keyframes fall-snow-windy {
                    0% { transform: translateY(-20px) translateX(0) rotate(0deg); }
                    100% { transform: translateY(105vh) translateX(65vw) rotate(720deg); }
                }
                @keyframes fall-blossom {
                    0% { transform: translateY(-20px) translateX(0) rotate(0deg); }
                    50% { transform: translateY(50vh) translateX(30px) rotate(120deg); }
                    100% { transform: translateY(105vh) translateX(-20px) rotate(240deg); }
                }
                @keyframes fall-blossom-windy {
                    0% { transform: translateY(-20px) translateX(0) rotate(0deg); }
                    100% { transform: translateY(105vh) translateX(75vw) rotate(540deg); }
                }
                @keyframes fall-leaves {
                    0% { transform: translateY(-20px) translateX(0) rotate(0deg); }
                    50% { transform: translateY(50vh) translateX(-20px) rotate(180deg); }
                    100% { transform: translateY(105vh) translateX(15px) rotate(360deg); }
                }
                @keyframes fall-leaves-windy {
                    0% { transform: translateY(-20px) translateX(0) rotate(0deg); }
                    100% { transform: translateY(105vh) translateX(70vw) rotate(480deg); }
                }
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
    },
    
    handleBgm(enabled, url) {
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
                audio.onended = () => {
                    let currentIndex = parseInt(audio.getAttribute('data-current-index') || "0", 10);
                    let nextIndex = (currentIndex + 1) % playlist.length;
                    audio.setAttribute('data-current-index', String(nextIndex));
                    audio.src = playlist[nextIndex];
                    
                    // [호환성 방어] play() 반환 Promise가 undefined일 수 있는 구형 모바일 브라우저 예외 처리
                    const playPromise = audio.play();
                    if (playPromise !== undefined) {
                        playPromise.then(() => {
                            console.log("Playing next BGM track:", playlist[nextIndex]);
                        }).catch(e => console.log("Next track play failed:", e));
                    }
                };
            } else {
                audio.loop = true;
                audio.onended = null;
            }
        }
        
        let btn = document.getElementById('village-bgm-toggle');
        if (!enabled) {
            // [원장님 피드백 반영] BGM이 비활성화 상태라면 화면에서 음악 버튼을 완전히 숨기고 즉시 정지!
            if (btn) btn.style.display = 'none';
            audio.pause();
            return;
        }
        
        // BGM이 다시 활성화되면 음악 버튼을 노출
        if (btn) {
            btn.style.display = 'flex';
        } else {
            btn = document.createElement('div');
            btn.id = 'village-bgm-toggle';
            btn.style = "position:fixed; bottom:145px; right:15px; width:45px; height:45px; border-radius:50%; background:rgba(255,255,255,0.85); box-shadow:0 8px 32px rgba(31,38,135,0.15); backdrop-filter:blur(6px); border:1.5px solid rgba(255,255,255,0.18); display:flex; align-items:center; justify-content:center; cursor:pointer; z-index:9998; transition:transform 0.3s ease;";
            btn.innerHTML = `<i class="fas fa-music" style="color:var(--v-wood); font-size:1.1rem; transition: color 0.3s;"></i>`;
            document.body.appendChild(btn);
            
            btn.onclick = () => {
                if (audio.paused) {
                    audio.play().then(() => {
                        btn.querySelector('i').style.color = '#e74c3c';
                        btn.style.transform = 'scale(1.1) rotate(15deg)';
                        localStorage.setItem('v44_bgm_user_play', 'true');
                    }).catch(e => console.log("BGM Play error:", e));
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
            const startPlay = () => {
                audio.play().then(() => {
                    btn.querySelector('i').style.color = '#e74c3c';
                    btn.style.transform = 'scale(1.1) rotate(15deg)';
                    document.removeEventListener('click', startPlay);
                    document.removeEventListener('touchstart', startPlay);
                }).catch(e => console.log("BGM play failed on interaction:", e));
            };
            
            audio.play().then(() => {
                btn.querySelector('i').style.color = '#e74c3c';
                btn.style.transform = 'scale(1.1) rotate(15deg)';
            }).catch(e => {
                console.log("Auto-play blocked by browser; registering interaction fallback.");
                document.addEventListener('click', startPlay);
                document.addEventListener('touchstart', startPlay);
            });
        } else {
            audio.pause();
            btn.querySelector('i').style.color = 'var(--v-wood)';
            btn.style.transform = 'scale(1) rotate(0deg)';
        }
    },
    
    confirmRankingWarp(event) {
        if (event) event.stopPropagation();
        
        const params = new URLSearchParams(window.location.search);
        let phone = (params.get('phone') || '').trim();
        if (!phone) phone = (localStorage.getItem('v44_user_phone') || '').trim();
        
        showAppConfirm("🏆 위대한 모험가들의 공간 [명예의 전당]으로 이동하시겠습니까?", () => {
            navigateTo('halloffame', { phone: phone });
        }, "🏆");
    },
    
    warpToRanking(event) {
        if (event) event.stopPropagation();
        const params = new URLSearchParams(window.location.search);
        let phone = (params.get('phone') || '').trim();
        if (!phone) phone = (localStorage.getItem('v44_user_phone') || '').trim();
        navigateTo('halloffame', { phone: phone });
    },
    
    toggleWeatherParticles(event) {
        if (event) event.stopPropagation();
        
        // [v65.0] 마스터 환경 마법 일괄 제어 토글 처리
        const settings = this.lastSettings || {};
        const magicEnabled = settings.magicEnabled === undefined || settings.magicEnabled === null || settings.magicEnabled.toString().toLowerCase() !== 'false';
        if (!magicEnabled) {
            showAppAlert("🪄 현재 마을 마법 환경이 전원 꺼진 상태(일괄 정지)입니다. 생텀에서 [🪄 마을 마법 환경 일괄 활성화] 체크 후 저장해 주세요.", "error", "마법 제어 불가");
            return;
        }
        
        const current = localStorage.getItem('village_weather_disabled') === 'true';
        const newVal = !current;
        localStorage.setItem('village_weather_disabled', String(newVal));
        
        const btn = document.getElementById('jeju-weather-toggle-btn');
        if (btn) {
            if (newVal) {
                btn.innerHTML = '🚫 켜기';
                btn.classList.add('disabled');
            } else {
                btn.innerHTML = '👁️ 끄기';
                btn.classList.remove('disabled');
            }
        }
        
        const weather = newVal ? 'sun' : (settings.weather || 'sun');
        const windSpeed = newVal ? 0 : (parseFloat(settings.realJejuWind) || 0);
        this.renderWeatherParticles(weather, windSpeed);
    }
};

window.onload = () => Village.init();
window.syncClubRecord = () => Village.openQuestModal('sync');
window.triggerSuddenMission = () => Village.openQuestModal('bonus');

/**
 * 📜 [v58.4] 웰니스 로그 시스템 바텀시트 제어 및 연동 로직
 */
window.openWellnessLogModal = function() {
  const modal = document.getElementById('wellness-log-modal');
  if (!modal) return;
  modal.style.display = 'flex';

  const todayList = document.getElementById('today-log-list');
  const historyList = document.getElementById('history-log-list');

  // 로딩 상태 렌더링
  if (todayList) todayList.innerHTML = '<div style="text-align:center; padding:20px; font-size:0.75rem; color:#94a3b8; font-weight:800;">⏱️ 모험 기록 조회 중...</div>';
  if (historyList) historyList.innerHTML = '<div style="text-align:center; padding:20px; font-size:0.75rem; color:#94a3b8; font-weight:800;">⏱️ 히스토리 조회 중...</div>';

  // 탭 리셋
  window.switchWellnessTab('today');

  const params = new URLSearchParams(window.location.search);
  let phone = (params.get('phone') || '').trim();
  if (!phone) phone = (localStorage.getItem('v44_user_phone') || '').trim();
  if (!phone) {
    if (todayList) todayList.innerHTML = '<div style="text-align:center; padding:20px; font-size:0.75rem; color:#ef4444; font-weight:800;">🚨 연락처 정보가 없습니다.</div>';
    return;
  }

  if (typeof google !== 'undefined' && google.script && google.script.run) {
    google.script.run
      .withSuccessHandler(function(res) {
        if (res && res.success) {
          window.renderWellnessLogs(res.todayLogs, res.historyLogs);
        } else {
          window.showWellnessError(res.error || "일치하는 데이터가 없습니다.");
        }
      })
      .withFailureHandler(function(err) {
        window.showWellnessError(err.message || err);
      })
      .getUserWellnessActivityHistory(phone);
  } else {
    // 로컬 브릿지 테스트 지원
    if (window.googleScriptRunMock && window.googleScriptRunMock.getUserWellnessActivityHistory) {
      window.googleScriptRunMock.getUserWellnessActivityHistory(phone);
    } else {
      window.showWellnessError("구글 서비스 연결 실패");
    }
  }
};

window.closeWellnessLogModal = function() {
  const modal = document.getElementById('wellness-log-modal');
  if (modal) modal.style.display = 'none';
};

window.switchWellnessTab = function(tab) {
  const todayBtn = document.getElementById('tab-today-log');
  const historyBtn = document.getElementById('tab-history-log');
  const todayContent = document.getElementById('content-today-log');
  const historyContent = document.getElementById('content-history-log');

  if (tab === 'today') {
    if (todayBtn) {
      todayBtn.style.background = '#fff';
      todayBtn.style.color = '#4f46e5';
      todayBtn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.05)';
    }
    if (historyBtn) {
      historyBtn.style.background = 'transparent';
      historyBtn.style.color = '#64748b';
      historyBtn.style.boxShadow = 'none';
    }
    if (todayContent) todayContent.style.display = 'block';
    if (historyContent) historyContent.style.display = 'none';
  } else {
    if (historyBtn) {
      historyBtn.style.background = '#fff';
      historyBtn.style.color = '#4f46e5';
      historyBtn.style.boxShadow = '0 2px 4px rgba(0,0,0,0.05)';
    }
    if (todayBtn) {
      todayBtn.style.background = 'transparent';
      todayBtn.style.color = '#64748b';
      todayBtn.style.boxShadow = 'none';
    }
    if (todayContent) todayContent.style.display = 'none';
    if (historyContent) historyContent.style.display = 'block';
  }
};

window.renderWellnessLogs = function(todayLogs, historyLogs) {
  const todayList = document.getElementById('today-log-list');
  const historyList = document.getElementById('history-log-list');
  const todayTotalScoreLog = document.getElementById('today-total-score-log');

  // 1. 오늘의 모험 로그 렌더링
  if (todayList) {
    if (!todayLogs || todayLogs.length === 0) {
      todayList.innerHTML = '<div style="text-align:center; padding:30px; font-size:0.75rem; color:#94a3b8; font-weight:800;">💤 오늘 달성한 모험 기록이 없습니다.</div>';
      if (todayTotalScoreLog) todayTotalScoreLog.innerText = '오늘 0점 획득';
    } else {
      let todayHtml = '';
      let calculatedTodayScore = 0;

      todayLogs.forEach(log => {
        // 이모지 추론 엔진
        let icon = '⚔️';
        if (log.indexOf('로그인') > -1) icon = '🔑';
        else if (log.indexOf('식단') > -1) icon = '🍱';
        else if (log.indexOf('티') > -1) icon = '🍵';
        else if (log.indexOf('스트레칭') > -1) icon = '🧘';
        else if (log.indexOf('방문') > -1 || log.indexOf('출석') > -1 || log.indexOf('퇴실') > -1) icon = '🏛️';
        else if (log.indexOf('운동강도') > -1) icon = '🤸‍♀️';
        else if (log.indexOf('워터') > -1 || log.indexOf('물') > -1) icon = '💧';
        else if (log.indexOf('칭찬') > -1) icon = '👏';
        else if (log.indexOf('보너스') > -1) icon = '✨';

        // 점수 숫자 추출 (🔒 7000보 버그 방지: 괄호 안의 획득 점수만 고속 정밀 파싱)
        let pointsMatch = log.match(/\((\d+)\)/);
        let points = pointsMatch ? Number(pointsMatch[1]) : 0;
        calculatedTodayScore += points;

        // 깔끔하게 점수 형태 정제
        let cleanText = log;
        cleanText = cleanText.replace(/\(\d+\)/g, '').trim();

        todayHtml += `
          <div style="display:flex; justify-content:space-between; align-items:center; padding:10px 14px; background:#f8fafc; border-radius:12px; border:1px solid #f1f5f9;">
            <div style="display:flex; align-items:center; gap:8px;">
              <span style="font-size:0.95rem;">${icon}</span>
              <span style="font-size:0.75rem; font-weight:800; color:#334155;">${cleanText}</span>
            </div>
            <span style="font-size:0.72rem; font-weight:900; color:#4f46e5; background:rgba(79, 70, 229, 0.06); padding:2px 8px; border-radius:8px;">+${points} 점</span>
          </div>
        `;
      });
      todayList.innerHTML = todayHtml;
      if (todayTotalScoreLog) todayTotalScoreLog.innerText = `오늘 총 +${calculatedTodayScore}점 획득 🔥`;
    }
  }

  // 2. 일별 히스토리 렌더링
  if (historyList) {
    if (!historyLogs || historyLogs.length === 0) {
      historyList.innerHTML = '<div style="text-align:center; padding:30px; font-size:0.75rem; color:#94a3b8; font-weight:800;">📈 과거 기록이 존재하지 않습니다.</div>';
    } else {
      let historyHtml = `
        <table style="width:100%; border-collapse:collapse; font-size:0.75rem; text-align:center;">
          <thead>
            <tr style="background:#f8fafc; border-bottom:1.5px solid #edf2f7; color:#64748b; font-weight:900;">
              <th style="padding:10px 4px; font-size:0.68rem;">모험 날짜</th>
              <th style="padding:10px 4px; font-size:0.68rem;">하루 획득 총점</th>
              <th style="padding:10px 4px; font-size:0.68rem;">달성 상태</th>
            </tr>
          </thead>
          <tbody>
      `;

      historyLogs.forEach((h, idx) => {
        let rowBg = idx % 2 === 0 ? '#ffffff' : '#f8fafc';
        
        let badge = '';
        if (h.score >= 120) badge = '<span style="background:#fef3c7; color:#d97706; font-weight:900; padding:2px 6px; border-radius:6px; font-size:0.62rem;">👑 전설적</span>';
        else if (h.score >= 80) badge = '<span style="background:#dcfce7; color:#15803d; font-weight:900; padding:2px 6px; border-radius:6px; font-size:0.62rem;">🔥 대활약</span>';
        else badge = '<span style="background:#f1f5f9; color:#64748b; font-weight:800; padding:2px 6px; border-radius:6px; font-size:0.62rem;">🌱 달성</span>';

        historyHtml += `
          <tr style="background:${rowBg}; border-bottom:1px solid #edf2f7; font-weight:800; color:#334155;">
            <td style="padding:11px 4px; font-weight:900; color:#1a202c; font-size:0.72rem;">${h.date}</td>
            <td style="padding:11px 4px; font-weight:900; color:#4f46e5; font-size:0.75rem;">+${h.score} 점</td>
            <td style="padding:11px 4px;">${badge}</td>
          </tr>
        `;
      });

      historyHtml += '</tbody></table>';
      historyList.innerHTML = historyHtml;
    }
  }
};

window.showWellnessError = function(msg) {
  const todayList = document.getElementById('today-log-list');
  const historyList = document.getElementById('history-log-list');
  if (todayList) todayList.innerHTML = `<div style="text-align:center; padding:20px; font-size:0.75rem; color:#ef4444; font-weight:800;">🚨 에러: ${msg}</div>`;
  if (historyList) historyList.innerHTML = `<div style="text-align:center; padding:20px; font-size:0.75rem; color:#ef4444; font-weight:800;">🚨 에러: ${msg}</div>`;
};
