/**
 * Nohyung Village Dashboard Logic (v44.3 - Premium Aesthetic Restoration)
 * Features: High-End Floating Nav, 3D Interactive Buttons, Integrated Evolution Flow
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

    rankings: [
        { type: "금주 랭킹", content: "체력왕: 홍길동 | 수행왕: 김개똥 | 수호왕: 이성실" },
        { type: "월간 랭킹", content: "다이어트킹: 박지니 | 미션킹: 최열정" },
        { type: "토탈 랭킹", content: "1위: 전설모험가 (120k) | 2위: 꾸준지존 (115k)" }
    ],
    currentRankIndex: 0,

    init() {
        console.log("Nohyung Village v44.3: Premium Restoration Initialized.");
        this.renderAll();
        this.updateEvolution();
        this.startTicker();
        this.bindEvents();
    },

    startTicker() {
        const ticker = document.getElementById('ranking-ticker');
        setInterval(() => {
            this.currentRankIndex = (this.currentRankIndex + 1) % this.rankings.length;
            const r = this.rankings[this.currentRankIndex];
            ticker.style.opacity = 0;
            setTimeout(() => {
                ticker.innerHTML = `<span style="font-size:0.7rem; color:var(--text-dim); display:block; font-weight:800;">[${r.type}]</span>${r.content}`;
                ticker.style.opacity = 1;
                ticker.style.transition = 'opacity 0.6s ease';
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
            document.getElementById('evo-bar').style.width = `${progress}%`;
            document.getElementById('evo-percent').innerText = `${Math.round(progress)}%`;
        }
    },

    setPerspective(view) {
        this.perspective = view;
        const bW = document.getElementById('btn-weekly');
        const bM = document.getElementById('btn-monthly');
        
        if (view === 'weekly') {
            bW.classList.add('active'); bM.classList.remove('active');
        } else {
            bM.classList.add('active'); bW.classList.remove('active');
        }
        
        this.renderAll();
    },

    renderHabits() {
        const container = document.getElementById('habit-list-container');
        if (!container) return;
        container.innerHTML = this.user.habits.map(h => `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:18px; background:${h.id === 'plus' ? 'rgba(241, 196, 15, 0.1)' : '#f9f9f9'}; border-radius:20px; margin-bottom:12px; border:2px solid ${h.id === 'plus' ? 'var(--gold)' : 'transparent'}; box-shadow: 0 4px 10px rgba(0,0,0,0.03);">
                <span style="font-size:1.05rem; font-weight:900; color:var(--text-main);">${h.title}</span>
                <div onclick="Village.checkHabit('${h.id}')" style="width:36px; height:36px; border:3.5px solid var(--v-border); border-radius:12px; cursor:pointer; display:flex; align-items:center; justify-content:center; background:${h.done ? 'var(--emerald)' : 'transparent'}; transition:all 0.2s;">
                    ${h.done ? '<i class="fa-solid fa-check" style="color:#fff;"></i>' : ''}
                </div>
            </div>
        `).join('');
    },

    checkHabit(id) {
        const habit = this.user.habits.find(h => h.id === id);
        if (habit && !habit.done) {
            const ok = confirm(`🏡 [습관 수호 성공!] 기록소(미라클 아카이브)로 이동하여 인증하시겠습니까?`);
            habit.done = true;
            this.user.totalScore += 10;
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
        alert("⚡ [계시: 돌발 미션 선포!]\n이장님의 긴급 퀘스트가 하사되었습니다! 마을 공지를 확인하세요!");
    },

    async syncClubRecord() {
        setTimeout(() => {
            alert("🏡 [클럽 동기화 완료!] 성장의 기록이 마을 역사에 새겨졌습니다.");
            this.user.stats.weekly.perf += 50;
            this.renderAll();
            this.updateEvolution();
        }, 1200);
    }
};

window.onload = () => Village.init();
window.syncClubRecord = () => Village.syncClubRecord();
window.triggerSuddenMission = () => Village.triggerSuddenMission();
