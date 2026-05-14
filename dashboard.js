/**
 * Nohyung Village Dashboard Logic (v42.1 - Refined Landmarks)
 * Features: Pillar of Records (Ranking), Village Wall (Notice), Sudden Alert
 */

const Village = {
    user: {
        name: "모험가",
        tier: "꿈나무 요정",
        stats: {
            health: { total: 15400, weekly: 450, monthly: 1200 },
            perf: { total: 24500, weekly: 680, monthly: 2400 },
            def: { total: 18200, weekly: 320, monthly: 950 }
        },
        water: 1.2,
        habits: [
            { id: 'h1', title: '모닝 티', guide: '기상 직후 따뜻한 물 한잔으로 대사 점화!', score: 2, done: false },
            { id: 'h2', title: '베지 퍼스트', guide: '혈당 관리를 위해 채소를 먼저 드세요.', score: 2, done: false },
            { id: 'h3', title: '슬로우 치잉', guide: '20회 이상 꼭꼭 씹어 소화를 돕습니다.', score: 2, done: false },
            { id: 'h4', title: '일일 7,000보', guide: '기초 대사량을 지키는 든든한 습관!', score: 3, done: false },
            { id: 'h5', title: '계단 마법', guide: '하체 근육은 제 2의 심장입니다.', score: 2, done: false },
            { id: 'h6', title: '나이트 컷', guide: '20시 이후 금식은 지방 연소의 지름길.', score: 10, done: false },
            { id: 'h7', title: '굿 슬립', guide: '성장 호르몬을 위한 자정 전 취침.', score: 2, done: false },
            { id: 'h8', title: '셀프 칭찬', guide: '나를 아끼는 마음이 성장의 동력입니다.', score: 2, done: false },
            { id: 'h9', title: '스트레칭', guide: '유연한 몸에 건강한 정신이 깃듭니다.', score: 2, done: false }
        ]
    },

    rankings: [
        { category: "지난주 킹", name: "강열정 님" },
        { category: "이번 달 킹", name: "이성실 님" },
        { category: "전설의 수호자", name: "박지니 님" }
    ],
    currentRankIndex: 0,

    init() {
        console.log("Welcome to Nohyung Village v42.1 (Pillar & Wall Mode)!");
        this.renderStatus();
        this.renderHabits();
        this.startPillarRotation();
        this.bindEvents();
    },

    // 🏛️ Pillar Rotation (Ranking Billboard)
    startPillarRotation() {
        const pillarText = document.getElementById('ranking-text');
        setInterval(() => {
            this.currentRankIndex = (this.currentRankIndex + 1) % this.rankings.length;
            const r = this.rankings[this.currentRankIndex];
            pillarText.style.opacity = 0;
            setTimeout(() => {
                pillarText.innerHTML = `[${r.category}]<br>${r.name}`;
                pillarText.style.opacity = 1;
                pillarText.style.transition = 'opacity 0.5s';
            }, 500);
        }, 5000);
    },

    // ⚡ Sudden Mission Logic
    triggerSuddenMission() {
        const missions = [
            { title: "지금 바로 플랭크!", desc: "60초 플랭크 인증 시 방어력 +20 EXP 하사!" },
            { title: "깜짝 수분 보충!", desc: "지금 500ml 원샷 후 슬라이더 밀면 +10 EXP!" },
            { title: "동료 칭찬하기", desc: "여행자 숙소에 따뜻한 댓글 1개 작성 시 +5 EXP!" }
        ];
        const m = missions[Math.floor(Math.random() * missions.length)];
        document.getElementById('sudden-title').innerText = m.title;
        document.getElementById('sudden-desc').innerText = m.desc;
        document.getElementById('sudden-alert').classList.add('active');
    },

    hideSuddenMission() {
        document.getElementById('sudden-alert').classList.remove('active');
    },

    renderStatus() {
        document.getElementById('user-name').innerText = this.user.name;
        document.getElementById('tier-name').innerText = this.user.tier;
        document.getElementById('health-total').innerText = this.user.stats.health.total.toLocaleString();
        document.getElementById('health-weekly').innerText = this.user.stats.health.weekly.toLocaleString();
        document.getElementById('perf-total').innerText = this.user.stats.perf.total.toLocaleString();
        document.getElementById('perf-weekly').innerText = this.user.stats.perf.weekly.toLocaleString();
        document.getElementById('def-total').innerText = this.user.stats.def.total.toLocaleString();
        document.getElementById('def-weekly').innerText = this.user.stats.def.weekly.toLocaleString();
        document.getElementById('water-val').innerText = `${this.user.water}L / 2.0L`;
    },

    renderHabits() {
        const container = document.getElementById('habit-list-container');
        if (!container) return;
        container.innerHTML = this.user.habits.map(h => `
            <div class="habit-row" title="${h.guide}" onclick="Village.showHabitGuide('${h.id}')" style="cursor: pointer; padding: 10px; background: rgba(255,255,255,0.3); border-radius: 12px; margin-bottom: 5px; display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size: 0.8rem; font-weight: 700;">${h.title}</span>
                <div style="width: 20px; height: 20px; border: 2px solid var(--village-border); border-radius: 4px; display:flex; align-items:center; justify-content:center;">
                    ${h.done ? '✅' : ''}
                </div>
            </div>
        `).join('');
    },

    showHabitGuide(id) {
        const habit = this.user.habits.find(h => h.id === id);
        this.showVillageNotice(`[습관 비책: ${habit.title}]`, `${habit.guide}\n\n완수 시 방어력 +${habit.score} EXP!`);
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

    async syncClubRecord() {
        setTimeout(() => {
            this.showVillageNotice("이장님의 전갈", "오늘의 클럽 활동이 동기화되었습니다! \n수행력(Perf) 점수가 상승했습니다. ✨");
            this.user.stats.perf.total += 35;
            this.renderStatus();
        }, 1200);
    },

    showVillageNotice(title, msg) {
        alert(`🏡 [노형 빌리지 소식] 🏡\n\n${title}\n${msg}`);
    }
};

// Start Village
window.onload = () => Village.init();
window.syncClubRecord = () => Village.syncClubRecord();
window.triggerSuddenMission = () => Village.triggerSuddenMission();
window.hideSuddenMission = () => Village.hideSuddenMission();
