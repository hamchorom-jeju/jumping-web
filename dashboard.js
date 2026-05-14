/**
 * Nohyung Village Dashboard Logic (v42.0 - Warm Game Fantasy)
 * Concept: Status Window (Health, Performance, Defense) + Quest Center
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
            { id: 'h1', title: '모닝 티', guide: '기상 직후 따뜻한 물 한잔으로 대사를 점화하세요!', score: 2, done: false },
            { id: 'h2', title: '베지 퍼스트', guide: '혈당 스파이크 방지를 위해 채소를 먼저 섭취하세요.', score: 2, done: false },
            { id: 'h3', title: '슬로우 치잉', guide: '한 입당 20회 이상 천천히 저작하여 소화를 돕습니다.', score: 2, done: false },
            { id: 'h4', title: '일일 7,000보', guide: '꾸준한 활동량 확보는 기초대사량 유지의 핵심입니다.', score: 3, done: false },
            { id: 'h5', title: '계단 마법', guide: '3층 이하는 계단을 이용해 하체 근력을 키우세요!', score: 2, done: false },
            { id: 'h6', title: '나이트 컷', guide: '20시 이후 금식은 지방 연소를 위한 강력한 규칙입니다.', score: 10, done: false },
            { id: 'h7', title: '굿 슬립', guide: '자정 전 취침은 세포 재생과 호복을 최적화합니다.', score: 2, done: false },
            { id: 'h8', title: '셀프 칭찬', guide: '노력한 자신을 위한 한마디는 지속 가능한 힘이 됩니다.', score: 2, done: false },
            { id: 'h9', title: '스트레칭', guide: '5분의 스트레칭이 하루의 컨디션을 결정합니다.', score: 2, done: false }
        ]
    },

    init() {
        console.log("Welcome to Nohyung Village v42.0!");
        this.renderStatus();
        this.renderHabits();
        this.bindEvents();
    },

    renderStatus() {
        document.getElementById('user-name').innerText = this.user.name;
        document.getElementById('tier-name').innerText = this.user.tier;
        
        // Health
        document.getElementById('health-total').innerText = this.user.stats.health.total.toLocaleString();
        document.getElementById('health-weekly').innerText = this.user.stats.health.weekly.toLocaleString();
        document.getElementById('health-monthly').innerText = this.user.stats.health.monthly.toLocaleString();
        
        // Performance
        document.getElementById('perf-total').innerText = this.user.stats.perf.total.toLocaleString();
        document.getElementById('perf-weekly').innerText = this.user.stats.perf.weekly.toLocaleString();
        document.getElementById('perf-monthly').innerText = this.user.stats.perf.monthly.toLocaleString();
        
        // Defense
        document.getElementById('def-total').innerText = this.user.stats.def.total.toLocaleString();
        document.getElementById('def-weekly').innerText = this.user.stats.def.weekly.toLocaleString();
        document.getElementById('def-monthly').innerText = this.user.stats.def.monthly.toLocaleString();

        document.getElementById('water-val').innerText = `${this.user.water}L / 2.0L`;
    },

    renderHabits() {
        const container = document.getElementById('habit-list-container');
        if (!container) return;

        container.innerHTML = this.user.habits.map(h => `
            <div class="habit-row" title="${h.guide} (점수: +${h.score} EXP)">
                <span class="habit-name" onclick="Village.showHabitGuide('${h.id}')">${h.title} <i class="fa-solid fa-circle-info" style="font-size:0.6rem; color:var(--text-dim);"></i></span>
                <div class="habit-check" onclick="Village.checkHabit('${h.id}')" style="background: ${h.done ? 'var(--defense)' : 'transparent'}; color: white;">
                    ${h.done ? '<i class="fa-solid fa-check"></i>' : ''}
                </div>
            </div>
        `).join('');
    },

    showHabitGuide(id) {
        const habit = this.user.habits.find(h => h.id === id);
        this.showVillageNotice(`[습관 비책: ${habit.title}]`, `${habit.guide}\n\n완수 시 방어력 +${habit.score} EXP 하사!`);
    },

    checkHabit(id) {
        const habit = this.user.habits.find(h => h.id === id);
        if (habit && !habit.done) {
            habit.done = true;
            this.user.stats.def.total += habit.score;
            this.user.stats.def.weekly += habit.score;
            this.renderStatus();
            this.renderHabits();
            this.showVillageNotice("방어력 상승!", `${habit.title} 완수! 습관이 당신을 지키는 방어력이 되었습니다. ✨`);
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

    async syncClubRecord() {
        console.log("Synchronizing with Nohyung Village Records...");
        setTimeout(() => {
            this.showVillageNotice("이장님의 전갈", "허허, 오늘 클럽 출석과 운동 기록이 확인되었소! \n수행력(미션) 점수가 크게 상승했구려! ✨");
            this.user.stats.perf.total += 35;
            this.user.stats.perf.weekly += 35;
            this.renderStatus();
        }, 1200);
    },

    showVillageNotice(title, msg) {
        alert(`🏡 [노형 빌리지 소식] 🏡\n\n${title}\n${msg}`);
    }
};

// Start Village Life
window.onload = () => Village.init();
window.syncClubRecord = () => Village.syncClubRecord();
