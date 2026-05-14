/**
 * Genie World Status Window Engine (v41.0)
 * Handles Quest Logic, Login Rewards, and Gamer-style UI interactions.
 */

const StatusWindow = {
    user: {
        name: "모험가",
        lifetimeExp: 84200,
        tier: "LEGENDARY JUMPER",
        monthlyPts: 4500,
        weeklyRank: 2,
        stats: { fat: 12.5, muscle: 34.2, energy: 1200 },
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

    init() {
        console.log("System: Status Window Activated...");
        this.renderAll();
        this.bindEvents();
        this.processLoginReward();
    },

    renderAll() {
        document.getElementById('user-name').innerText = this.user.name;
        document.getElementById('tier-name').innerText = this.user.tier;
        document.getElementById('lifetime-exp-val').innerText = this.user.lifetimeExp.toLocaleString();
        document.getElementById('monthly-pts').innerText = this.user.monthlyPts.toLocaleString();
        document.getElementById('weekly-rank').innerHTML = `${this.user.weeklyRank}<span class="unit">위</span>`;
        
        // Stats
        document.getElementById('stat-fat').innerText = this.user.stats.fat;
        document.getElementById('stat-muscle').innerText = this.user.stats.muscle;
        document.getElementById('stat-energy').innerText = this.user.stats.energy;

        // Progress Bar
        const progress = (this.user.lifetimeExp % 1000) / 10;
        document.getElementById('lifetime-exp-bar').style.width = `${progress}%`;
    },

    bindEvents() {
        const waterSlider = document.getElementById('water-range');
        if (waterSlider) {
            waterSlider.addEventListener('input', (e) => {
                const val = (e.target.value / 10).toFixed(1);
                document.getElementById('water-val').innerText = `${val}L / 2.0L`;
                this.updateEnergy(val);
            });
        }
    },

    updateEnergy(val) {
        this.user.stats.energy = Math.floor(val * 1000);
        document.getElementById('stat-energy').innerText = this.user.stats.energy;
    },

    processLoginReward() {
        // Immediate login reward visualization
        setTimeout(() => {
            this.showSystemMessage("입장 완료!", "오늘의 출석 보상으로 [+5 EXP]가 지급되었습니다. \n상태창이 갱신되었습니다.");
            this.user.lifetimeExp += 5;
            this.renderAll();
        }, 1000);
    },

    toggleHabits() {
        const habitCount = this.user.habits.filter(h => h.done).length;
        this.showSystemMessage("데일리 습관 퀘스트", `현재 완수 현황: ${habitCount} / 9 \n\n기록소에서 인증샷을 찍어 퀘스트를 완수하세요!`);
        // In actual implementation, this could open a modal or redirect
    },

    showSystemMessage(title, msg) {
        // Gamer-style alert popup
        const alertBox = document.createElement('div');
        alertBox.style.cssText = `
            position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%);
            background: rgba(255, 255, 255, 0.95); border: 2px solid #f1c40f;
            padding: 25px; border-radius: 20px; box-shadow: 0 20px 60px rgba(0,0,0,0.2);
            z-index: 9999; width: 300px; text-align: center; backdrop-filter: blur(10px);
        `;
        alertBox.innerHTML = `
            <h3 style="color: #f1c40f; margin-bottom: 15px; font-weight: 900;">[ SYSTEM ]</h3>
            <p style="font-weight: 800; font-size: 0.95rem; margin-bottom: 10px;">${title}</p>
            <p style="font-size: 0.8rem; color: #555; line-height: 1.5;">${msg}</p>
            <button onclick="this.parentElement.remove()" style="margin-top: 20px; padding: 10px 30px; border: none; background: #2ecc71; color: #fff; border-radius: 10px; font-weight: 800; cursor: pointer;">확인</button>
        `;
        document.body.appendChild(alertBox);
    }
};

// Start System
window.onload = () => StatusWindow.init();
window.toggleHabits = () => StatusWindow.toggleHabits();
