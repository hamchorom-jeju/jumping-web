/**
 * Nohyung Village Dashboard Logic (v44.5 - Premium Modal & Rich Content)
 * Features: Custom Achievement Modal, Deep Habit Guides, Stable v44.0 Balance
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
            { id: 'h1', title: '모닝 티', done: false, icon: '🍵', guide: '기상 후 따뜻한 물 한잔은 잠든 몸의 마력을 깨우고 신진대사의 길을 엽니다.' },
            { id: 'h2', title: '베지 퍼스트', done: false, icon: '🥗', guide: '채소를 먼저 섭취하여 혈당의 요동을 막고 몸의 균형을 수호하세요.' },
            { id: 'h3', title: '슬로우 치잉', done: false, icon: '⏳', guide: '천천히 꼭꼭 씹는 행위는 뇌에 포만감의 신호를 보내 과식을 막는 지혜입니다.' },
            { id: 'h4', title: '일일 7,000보', done: false, icon: '👣', guide: '꾸준한 발걸음은 대지의 기운을 흡수하고 심장을 강화하는 가장 정직한 수련입니다.' },
            { id: 'h5', title: '계단 마법', done: false, icon: '🪜', guide: '계단을 오르는 매 순간, 당신의 하체 근육은 더욱 견고한 성벽이 됩니다.' },
            { id: 'h6', title: '나이트 컷', done: false, icon: '🌙', guide: '20시 이후의 금식은 당신의 장기가 밤새 평온하게 휴식하고 회복하게 돕습니다.' },
            { id: 'h7', title: '굿 슬립', done: false, icon: '😴', guide: '자정 전 깊은 잠은 모험가에게 필요한 치유의 마법이 가장 활발히 일어나는 시간입니다.' },
            { id: 'h8', title: '셀프 칭찬', done: false, icon: '💖', guide: '스스로를 격려하는 따뜻한 말 한마디는 내면의 성장을 이끄는 가장 큰 동력입니다.' },
            { id: 'h9', title: '스트레칭', done: false, icon: '🧘', guide: '경직된 근육을 펴주는 것은 몸의 흐름을 원활하게 하여 부상을 막는 방어의 시작입니다.' },
            { id: 'plus', title: '✨ 미라클 플러스', done: false, icon: '🌟', guide: '새벽기상, 독서 등 인생을 풍요롭게 만드는 사소하지만 위대한 승리들을 기록하세요.' }
        ]
    },

    init() {
        console.log("Welcome to Nohyung Village v44.5: Premium Achievement System.");
        this.renderAll();
        this.updateEvolution();
        this.startTicker();
        this.bindEvents();
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
            <div style="display:flex; justify-content:space-between; align-items:center; padding:15px; background:${h.id === 'plus' ? '#fff9c4' : '#fff'}; border-radius:20px; margin-bottom:10px; border:2px solid ${h.id === 'plus' ? 'var(--def)' : 'transparent'}; box-shadow:0 4px 10px rgba(0,0,0,0.03);">
                <span style="font-size:1rem; font-weight:900; color:var(--v-wood);">${h.title}</span>
                <div onclick="Village.openModal('${h.id}')" style="width:34px; height:34px; border:3px solid var(--v-border); border-radius:10px; cursor:pointer; display:flex; align-items:center; justify-content:center; background:${h.done ? 'var(--emerald)' : 'transparent'};">
                    ${h.done ? '✅' : ''}
                </div>
            </div>
        `).join('');
    },

    // ✨ Premium Modal Logic
    openModal(id) {
        const habit = this.user.habits.find(h => h.id === id);
        if (!habit) return;

        document.getElementById('modal-habit-icon').innerText = habit.icon;
        document.getElementById('modal-habit-title').innerText = habit.title;
        document.getElementById('modal-habit-guide').innerText = habit.guide;
        
        const confirmBtn = document.getElementById('modal-confirm-btn');
        confirmBtn.onclick = () => {
            this.confirmHabit(id);
        };

        document.getElementById('habit-modal').style.display = 'flex';
    },

    closeModal() {
        document.getElementById('habit-modal').style.display = 'none';
    },

    confirmHabit(id) {
        const habit = this.user.habits.find(h => h.id === id);
        if (habit && !habit.done) {
            habit.done = true;
            this.user.totalScore += 10;
            this.user.stats.weekly.def += 10;
            this.renderAll();
            this.updateEvolution();
        }
        this.closeModal();
        location.href = `miracle.html?cat=${(id === 'plus' ? 'plus' : 'habit')}&item=${id}`;
    },

    startTicker() {
        const ticker = document.getElementById('ranking-ticker');
        let idx = 0;
        const ranks = ["금주 체력왕: 홍길동 님 ✨", "금주 수행왕: 김개똥 님 🗡️", "금주 수호왕: 이성실 님 🛡️"];
        setInterval(() => {
            idx = (idx + 1) % ranks.length;
            ticker.style.opacity = 0;
            setTimeout(() => {
                ticker.innerText = ranks[idx];
                ticker.style.opacity = 1;
                ticker.style.transition = 'opacity 0.5s';
            }, 500);
        }, 5000);
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

    triggerSuddenMission() { alert("⚡ [돌발 미션 선포!]"); },
    async syncClubRecord() {
        setTimeout(() => {
            alert("🏡 [클럽 동기화 완료!]");
            this.user.stats.weekly.perf += 50;
            this.renderAll();
            this.updateEvolution();
        }, 1200);
    }
};

window.onload = () => Village.init();
window.syncClubRecord = () => Village.syncClubRecord();
window.triggerSuddenMission = () => Village.triggerSuddenMission();
