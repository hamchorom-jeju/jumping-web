/**
 * Nohyung Village Dashboard Logic (v44.93 - Logic Correction)
 * Features: Single/Dual Modal Buttons, Descriptive Guides, v44.0 Immutable Base
 */

const Village = {
    perspective: 'weekly',
    user: {
        name: "모험가",
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
            { name: "씨앗", min: 0 }, { name: "새싹", min: 1001 }, { name: "나무", min: 3001 },
            { name: "꽃", min: 8001 }, { name: "꿈나무 요정", min: 15001 },
            { name: "전설의 점퍼", min: 30001 }, { name: "지니 수호신", min: 60001 }
        ],
        water: 1.2,
        habits: [
            { id: 'h1', title: '모닝 티', done: false },
            { id: 'h2', title: '베지 퍼스트', done: false },
            { id: 'h3', title: '슬로우 치잉', done: false },
            { id: 'h4', title: '일일 7,000보', done: false },
            { id: 'h5', title: '계단 마법', done: false },
            { id: 'h6', title: '나이트 컷', done: false },
            { id: 'h7', title: '굿 슬립', done: false },
            { id: 'h8', title: '셀프 칭찬', done: false },
            { id: 'h9', title: '스트레칭', done: false },
            { id: 'plus', title: '✨ 미라클 플러스', done: false }
        ]
    },

    // 🏆 Multi-Tier Ranking Data
    rankings: [
        { type: "금주 랭킹", content: "체력왕: 홍길동 | 미션왕: 김개똥 | 수호왕: 이성실" },
        { type: "월간 랭킹", content: "다이어트킹: 박지니 | 미션 선두: 최열정" },
        { type: "토탈 랭킹", content: "1위: 전설모험가 | 2위: 꾸준지존 | 3위: 열정맨" }
    ],
    currentRankIndex: 0,

    // ⚔️ Quest Content Data (v44.96 Final Literal Refinement)
    quests: {
        sync: { 
            title: "클럽 동기화", icon: "⚡", 
            guide: "클럽 출석 점수를 반영하시겠어요?\n오늘 클럽에 출석 하셨다면 15포인트,\n운동량에 따라 최대 20포인트가 반영됩니다.", 
            btn: "동기화 수행", link: "#", single: true 
        },
        visit: { 
            title: "방문 인증", icon: "📸", 
            guide: "오늘 클럽에 출석하셨나요?\n그렇다면 방문 인증을 남겨보세요\n아카이브에 인증시 15점이 추가됩니다.", 
            btn: "인증하러 가기", link: "miracle.html?cat=visit", single: false 
        },
        meal: { 
            title: "식단 인증", icon: "🍱", 
            guide: "꾸준한 식단 기록은 강력한 변화의 열쇠입니다.\n아카이브에 인증시 최대 30점이 지급됩니다.", 
            btn: "인증하러 가기", link: "miracle.html?cat=meal", single: false 
        },
        water: { 
            title: "워터 헌터", icon: "💧", 
            guide: "수분 섭취량 만큼 게이지를 조정해보세요.\n섭취량에 따라 점수를 차등지급합니다. (최대 20점)\n아카이브에 인증시 15점이 추가됩니다.", 
            btn: "인증하러 가기", link: "miracle.html?cat=water", single: false 
        },
        bonus: { 
            title: "보너스 퀘스트", icon: "✨", 
            guide: "돌발 미션을 수행하시겠어요?\n아카이브로 이동해 인증을 남기실 수 있습니다.\n아카이브 인증시 15점이 추가됩니다.", 
            btn: "이동하기", link: "miracle.html?cat=bonus", single: false 
        }
    },

    init() {
        console.log("v44.93 Logical Correction Engine Initialized.");
        this.renderAll();
        this.updateEvolution();
        this.startTicker();
        this.bindEvents();
    },

    // ✨ Refined Modal Handling
    openQuestModal(key) {
        const q = this.quests[key];
        if (!q) return;
        
        document.getElementById('modal-habit-icon').innerText = q.icon;
        document.getElementById('modal-habit-title').innerText = q.title;
        document.getElementById('modal-habit-guide').innerText = q.guide;
        
        const cancelBtn = document.getElementById('modal-cancel-btn');
        const confirmBtn = document.getElementById('modal-confirm-btn');
        
        // Single/Dual Button Toggle
        if (q.single) {
            cancelBtn.style.display = 'none';
            confirmBtn.innerText = q.btn;
            confirmBtn.style.flex = "1";
        } else {
            cancelBtn.style.display = 'block';
            cancelBtn.innerText = "나중에";
            confirmBtn.innerText = q.btn;
            confirmBtn.style.flex = "1.5";
        }
        
        confirmBtn.onclick = () => {
            if (q.link !== "#") {
                location.href = q.link;
            } else if (key === 'sync') {
                this.syncClubRecordActual();
            }
            this.closeModal();
        };

        document.getElementById('habit-modal').style.display = 'flex';
    },

    closeModal() {
        document.getElementById('habit-modal').style.display = 'none';
    },

    syncClubRecordActual() {
        setTimeout(() => {
            alert("🏡 [클럽 동기화 완료!] 15 EXP가 반영되었습니다.");
            this.user.stats.weekly.perf += 15;
            this.user.totalScore += 15;
            this.renderAll();
            this.updateEvolution();
        }, 1000);
    },

    startTicker() {
        const ticker = document.getElementById('ranking-ticker');
        setInterval(() => {
            this.currentRankIndex = (this.currentRankIndex + 1) % this.rankings.length;
            const r = this.rankings[this.currentRankIndex];
            ticker.style.opacity = 0;
            setTimeout(() => {
                ticker.innerHTML = `<span style="font-size:0.75rem; opacity:0.7; display:block; margin-bottom:4px;">[${r.type}]</span>${r.content}`;
                ticker.style.opacity = 1;
                ticker.style.transition = 'opacity 0.5s';
            }, 500);
        }, 5000);
    },

    renderAll() {
        document.getElementById('user-name').innerText = this.user.name;
        document.getElementById('total-score').innerText = this.user.totalScore.toLocaleString();
        document.getElementById('current-rank').innerText = this.user.rank;
        document.getElementById('water-val').innerText = `${this.user.water}L / 2.0L`;
        const view = this.perspective;
        const currentData = this.user.stats[view];
        const maxData = this.user.max[view];
        this.updateGauge('health', currentData.health, maxData.health);
        this.updateGauge('perf', currentData.perf, maxData.perf);
        this.updateGauge('def', currentData.def, maxData.def);
        this.renderHabits();
    },

    updateGauge(id, val, max) {
        document.getElementById(`${id}-val`).innerText = val.toLocaleString();
        const percent = Math.min((val / max) * 100, 100);
        document.getElementById(`${id}-bar`).style.width = `${percent}%`;
        document.getElementById(`${id}-label`).innerText = `${this.perspective === 'weekly' ? '주간' : '월간'} 목표 대비: ${Math.round(percent)}%`;
    },

    updateEvolution() {
        const score = this.user.totalScore;
        let ci = 0;
        for (let i = 0; i < this.user.tiers.length; i++) {
            if (score >= this.user.tiers[i].min) ci = i;
        }
        const next = this.user.tiers[ci + 1];
        if (next) {
            const min = this.user.tiers[ci].min;
            const progress = ((score - min) / (next.min - min)) * 100;
            document.getElementById('evo-bar').style.width = `${Math.min(progress, 100)}%`;
            document.getElementById('evo-percent').innerText = `${Math.round(progress)}%`;
        }
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
                <div onclick="Village.checkHabit('${h.id}')" style="width:28px; height:28px; border:2px solid var(--v-border); border-radius:8px; cursor:pointer; display:flex; align-items:center; justify-content:center; background:${h.done ? 'var(--def)' : 'transparent'};">
                    ${h.done ? '✅' : ''}
                </div>
            </div>
        `).join('');
    },

    checkHabit(id) {
        const habit = this.user.habits.find(h => h.id === id);
        if (habit && !habit.done) {
            const ok = confirm(`🏡 [수호 완료!] 기록소로 이동하시겠습니까?`);
            habit.done = true;
            this.user.totalScore += 10;
            this.user.stats.weekly.def += 10;
            this.renderAll();
            this.updateEvolution();
            if (ok) location.href = `miracle.html?cat=${(id === 'plus' ? 'plus' : 'habit')}&item=${id}`;
        }
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

    triggerSuddenMission() {
        document.getElementById('sudden-mission-bar').style.display = 'block';
    }
};

window.onload = () => Village.init();
window.syncClubRecord = () => Village.openQuestModal('sync');
window.triggerSuddenMission = () => Village.openQuestModal('bonus');
