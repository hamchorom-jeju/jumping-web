/**
 * Nohyung Village Dashboard Logic (v44.0 - Perspective Gauges)
 * Features: Weekly/Monthly Toggle, Evolution Progress Bar, Context-Aware Stats
 */

const Village = {
    perspective: 'weekly', // 'weekly' or 'monthly'
    user: {
        name: "모험가",
        tier: "꿈나무 요정",
        totalScore: 84200,
        rank: 12,
        stats: {
            weekly: { health: 450, perf: 680, def: 320 },
            monthly: { health: 1200, perf: 2400, def: 950 }
        },
        // Maximum Potentials (Goal Posts)
        max: {
            weekly: { health: 600, perf: 1000, def: 700 },
            monthly: { health: 2400, perf: 4000, def: 2800 }
        },
        // Tier Thresholds for Evolution
        tiers: [
            { name: "씨앗", min: 0 },
            { name: "새싹", min: 1001 },
            { name: "나무", min: 3001 },
            { name: "꽃", min: 8001 },
            { name: "꿈나무 요정", min: 15001 },
            { name: "전설의 점퍼", min: 30001 },
            { name: "지니 수호신", min: 60001 }
        ],
        water: 1.2,
        habits: [
            { id: 'h1', title: '모닝 티', done: false, guide: '기상 후 따뜻한 물!' },
            { id: 'h2', title: '베지 퍼스트', done: false, guide: '채소 먼저 먹기!' },
            { id: 'h3', title: '슬로우 치잉', done: false, guide: '꼭꼭 씹기!' },
            { id: 'h4', title: '일일 7,000보', done: false, guide: '꾸준한 걷기!' },
            { id: 'h5', title: '계단 마법', done: false, guide: '계단 이용하기!' },
            { id: 'h6', title: '나이트 컷', done: false, guide: '20시 이후 금식!' },
            { id: 'h7', title: '굿 슬립', done: false, guide: '자정 전 취침!' },
            { id: 'h8', title: '셀프 칭찬', done: false, guide: '나를 아껴주기!' },
            { id: 'h9', title: '스트레칭', done: false, guide: '몸 풀어주기!' },
            { id: 'plus', title: '✨ 미라클 플러스', done: false, guide: '인생의 승리 기록!' }
        ]
    },

    init() {
        console.log("Welcome to Nohyung Village v44.0 (Interactive Perspective)!");
        this.renderAll();
        this.updateEvolution();
        this.bindEvents();
    },

    renderAll() {
        // Basic Info
        document.getElementById('user-name').innerText = this.user.name;
        document.getElementById('total-score').innerText = this.user.totalScore.toLocaleString();
        document.getElementById('current-rank').innerText = this.user.rank;
        document.getElementById('water-val').innerText = `${this.user.water}L / 2.0L`;

        // Ability Gauges based on Perspective
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

    // ✨ Evolution Logic (Tier Progress)
    updateEvolution() {
        const score = this.user.totalScore;
        let currentTierIndex = 0;
        for (let i = 0; i < this.user.tiers.length; i++) {
            if (score >= this.user.tiers[i].min) {
                currentTierIndex = i;
            }
        }

        const nextTier = this.user.tiers[currentTierIndex + 1];
        if (nextTier) {
            const currentTierMin = this.user.tiers[currentTierIndex].min;
            const progress = ((score - currentTierMin) / (nextTier.min - currentTierMin)) * 100;
            const finalPercent = Math.min(Math.max(progress, 0), 100);
            
            document.getElementById('evo-bar').style.width = `${finalPercent}%`;
            document.getElementById('evo-percent').innerText = `${Math.round(finalPercent)}%`;
        } else {
            document.getElementById('evo-bar').style.width = '100%';
            document.getElementById('evo-percent').innerText = 'MAX';
        }
    },

    setPerspective(view) {
        this.perspective = view;
        
        // Toggle Button Styles
        const btnWeekly = document.getElementById('btn-weekly');
        const btnMonthly = document.getElementById('btn-monthly');
        
        if (view === 'weekly') {
            btnWeekly.style.background = '#fff'; btnWeekly.style.boxShadow = '0 2px 5px rgba(0,0,0,0.1)';
            btnMonthly.style.background = 'transparent'; btnMonthly.style.boxShadow = 'none';
        } else {
            btnMonthly.style.background = '#fff'; btnMonthly.style.boxShadow = '0 2px 5px rgba(0,0,0,0.1)';
            btnWeekly.style.background = 'transparent'; btnWeekly.style.boxShadow = 'none';
        }
        
        this.renderAll();
    },

    renderHabits() {
        const container = document.getElementById('habit-list-container');
        if (!container) return;
        container.innerHTML = this.user.habits.map(h => `
            <div style="display:flex; justify-content:space-between; align-items:center; padding:12px; background:${h.id === 'plus' ? '#fff9c4' : '#fff'}; border-radius:15px; margin-bottom:8px; border:2px solid ${h.id === 'plus' ? 'var(--def)' : 'var(--v-border)'};">
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
            const confirmMove = confirm(`🏡 [수호 완료!]\n\n'${habit.title}' 실천 완료!\n\n인증 기록소로 이동하시겠습니까?`);
            habit.done = true;
            this.user.totalScore += 10;
            this.user.stats.weekly.def += 10;
            this.user.stats.monthly.def += 10;
            this.renderAll();
            this.updateEvolution();
            if (confirmMove) {
                location.href = `miracle.html?cat=${(id === 'plus') ? 'plus' : 'habit'}&item=${id}`;
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
        alert("⚡ [돌발 미션!]\n마을 공지 아래에 새로운 긴급 미션이 하사되었습니다!");
    },

    async syncClubRecord() {
        setTimeout(() => {
            alert("🏡 [클럽 동기화 완료!]");
            this.user.stats.weekly.perf += 50;
            this.user.totalScore += 50;
            this.renderAll();
            this.updateEvolution();
        }, 1200);
    }
};

// Start Village
window.onload = () => Village.init();
window.syncClubRecord = () => Village.syncClubRecord();
window.triggerSuddenMission = () => Village.triggerSuddenMission();
