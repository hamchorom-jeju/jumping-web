/**
 * Nohyung Village Dashboard Logic (v43.0 - Silver-Friendly & Gauges)
 * Features: Left/Right Status Split, Ability Gauges, 3-Tier Ticker, Habit Pop-up
 */

const Village = {
    user: {
        name: "모험가",
        tier: "꿈나무 요정",
        totalScore: 84200,
        rank: 12,
        stats: {
            health: { val: 15400, weekly: 450, monthly: 1200, max: 20000 },
            perf: { val: 24500, weekly: 680, monthly: 2400, max: 30000 },
            def: { val: 18200, weekly: 320, monthly: 950, max: 25000 }
        },
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
            { id: 'h9', title: '스트레칭', done: false }
        ]
    },

    // 📢 Hall of Fame Content
    rankings: [
        { type: "금주 랭킹", content: "체력왕: 홍길동 | 수행왕: 김개똥 | 수호왕: 이성실" },
        { type: "월간 랭킹", content: "다이어트킹: 박지니 | 미션킹: 최열정" },
        { type: "토탈 랭킹", content: "1위: 전설모험가 (120k) | 2위: 꾸준지존 (115k)" }
    ],
    currentRankIndex: 0,

    init() {
        console.log("Welcome to Nohyung Village v43.0 (Silver-Friendly Mode)!");
        this.renderAll();
        this.startTicker();
        this.bindEvents();
    },

    renderAll() {
        // Left Profile
        document.getElementById('user-name').innerText = this.user.name;
        document.getElementById('total-score').innerText = this.user.totalScore.toLocaleString();
        document.getElementById('current-rank').innerText = this.user.rank;

        // Right Gauges
        this.updateGauge('health', this.user.stats.health);
        this.updateGauge('perf', this.user.stats.perf);
        this.updateGauge('def', this.user.stats.def);

        // Water & Habits
        document.getElementById('water-val').innerText = `${this.user.water}L / 2.0L`;
        this.renderHabits();
    },

    updateGauge(id, data) {
        document.getElementById(`${id}-val`).innerText = data.val.toLocaleString();
        document.getElementById(`${id}-weekly`).innerText = data.weekly.toLocaleString();
        document.getElementById(`${id}-monthly`).innerText = data.monthly.toLocaleString();
        
        const percent = Math.min((data.val / data.max) * 100, 100);
        document.getElementById(`${id}-bar`).style.width = `${percent}%`;
    },

    startTicker() {
        const ticker = document.getElementById('ranking-ticker');
        setInterval(() => {
            this.currentRankIndex = (this.currentRankIndex + 1) % this.rankings.length;
            const r = this.rankings[this.currentRankIndex];
            ticker.style.opacity = 0;
            setTimeout(() => {
                ticker.innerHTML = `<span style="font-size:0.7rem; color:#fff; display:block;">[${r.type}]</span>${r.content}`;
                ticker.style.opacity = 1;
                ticker.style.transition = 'opacity 0.5s';
            }, 500);
        }, 5000);
    },

    renderHabits() {
        const container = document.getElementById('habit-list-container');
        if (!container) return;
        container.innerHTML = this.user.habits.map(h => `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:12px; background:#fff; border-radius:15px; margin-bottom:8px; border:1px solid var(--v-border);">
                <span style="font-size:0.9rem; font-weight:800; color:var(--v-wood);">${h.title}</span>
                <div onclick="Village.checkHabit('${h.id}')" style="width:28px; height:28px; border:2px solid var(--v-border); border-radius:8px; cursor:pointer; display:flex; align-items:center; justify-content:center;">
                    ${h.done ? '✅' : ''}
                </div>
            </div>
        `).join('');
    },

    checkHabit(id) {
        const habit = this.user.habits.find(h => h.id === id);
        if (habit && !habit.done) {
            const confirmMove = confirm(`🏡 [습관 수호 완료!]\n\n'${habit.title}' 실천을 완료하셨군요!\n\n지금 바로 기록소(미라클 아카이브)로 이동하여\n인증샷을 남기고 추가 보너스를 받으시겠습니까?`);
            
            habit.done = true;
            this.user.totalScore += 10;
            this.user.stats.def.val += 10;
            this.renderAll();

            if (confirmMove) {
                location.href = `miracle.html?cat=habit&item=${id}`;
            }
        }
    },

    bindEvents() {
        const waterSlider = document.getElementById('water-range');
        if (waterSlider) {
            waterSlider.addEventListener('input', (e) => {
                const val = (e.target.value / 10).toFixed(1);
                document.getElementById('water-val').innerText = `${val}L / 2.0L`;
                this.user.water = parseFloat(val);
            });
        }
    },

    triggerSuddenMission() {
        const suddenBar = document.getElementById('sudden-mission-bar');
        suddenBar.style.display = 'block';
        alert("⚡ [이장님의 보너스 퀘스트!]\n\n마을 벽보 아래에 새로운 긴급 미션이 하사되었습니다!\n지금 바로 확인하고 수행하세요!");
    },

    async syncClubRecord() {
        setTimeout(() => {
            alert("🏡 [클럽 동기화 완료!]\n\n오늘의 방문 기록과 운동 데이터가 확인되었습니다.\n수행력(Perf) 점수가 상승했습니다!");
            this.user.stats.perf.val += 50;
            this.user.totalScore += 50;
            this.renderAll();
        }, 1200);
    }
};

// Start Village
window.onload = () => Village.init();
window.syncClubRecord = () => Village.syncClubRecord();
window.triggerSuddenMission = () => Village.triggerSuddenMission();
