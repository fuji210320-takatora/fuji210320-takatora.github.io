let fieldersData = [];
let pitchersData = [];
let draftPool = [];
let myTeam = { fielders: [], pitchers: [] };

let currentDraftType = 'fielder';
let currentCandidate = null;
let requiredCount = 9;
let skipsLeft = 5;

// ==========================================
// 1. 共通エクセルの読み込み
// ==========================================
window.addEventListener('DOMContentLoaded', async () => {
    try {
        // キャッシュを無視して最新のExcelを読み込む
        const res = await fetch('./baseball_data.xlsx?v=' + new Date().getTime());
        if (!res.ok) throw new Error("Fetch Error");
        
        const buffer = await res.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        
        const sheetF = workbook.Sheets['野手能力'];
        const sheetP = workbook.Sheets['投手能力'];
        
        const rawFielders = XLSX.utils.sheet_to_json(sheetF) || [];
        const rawPitchers = XLSX.utils.sheet_to_json(sheetP) || [];
        
        fieldersData = rawFielders.filter(p => p['選手名']);
        pitchersData = rawPitchers.filter(p => p['選手名']);
        
        document.getElementById('screen-loading').classList.remove('active');
        startDraft('fielder');
        
    } catch (e) {
        console.error("データの読み込みに失敗しました:", e);
        const loadingText = document.getElementById('loading-text');
        if (loadingText) loadingText.innerText = "データの読み込みに失敗しました。再読み込みしてください。";
    }
});

// ==========================================
// 2. 能力値のランク判定関数
// ==========================================
function getRankInfo(num) {
    if (isNaN(num)) return { rank: '-', color: '#999' };
    if (num >= 90) return { rank: 'S', color: '#D4AF37' };
    if (num >= 80) return { rank: 'A', color: '#F06292' };
    if (num >= 70) return { rank: 'B', color: '#E53935' };
    if (num >= 60) return { rank: 'C', color: '#FF9800' };
    if (num >= 50) return { rank: 'D', color: '#FBC02D' };
    if (num >= 40) return { rank: 'E', color: '#7CB342' };
    if (num >= 20) return { rank: 'F', color: '#03A9F4' };
    return { rank: 'G', color: '#9E9E9E' };
}

function createStatCard(label, val) {
    const num = parseInt(val, 10);
    const data = getRankInfo(num);
    const displayRank = data.rank === '-' ? '' : `<div class="stat-rank" style="color: ${data.color};">${data.rank}</div>`;
    
    return `
        <div class="stat-card">
            <div class="stat-label">${label}</div>
            ${displayRank}
            <div class="stat-val">${isNaN(num) ? '-' : num}</div>
        </div>
    `;
}

// ==========================================
// 3. UI更新・ドラフト進行
// ==========================================
function showScreen(id) {
    document.querySelectorAll('.screen').forEach(el => el.classList.remove('active'));
    document.getElementById(id).classList.add('active');
}

function updateDraftHeader() {
    const currentCount = currentDraftType === 'fielder' ? myTeam.fielders.length : myTeam.pitchers.length;
    const remain = requiredCount - currentCount;
    
    document.getElementById('draft-type-label').innerText = currentDraftType === 'fielder' ? "野手ドラフト" : "投手ドラフト";
    document.getElementById('draft-progress').innerHTML = `${currentCount} / ${requiredCount} <span style="font-size:14px; color:#666;">あと${remain}人</span>`;
    
    const skipBtn = document.getElementById('btn-skip');
    if (skipsLeft > 0) {
        document.getElementById('draft-skips').innerText = `見送り残り：${skipsLeft}回`;
        skipBtn.classList.remove('btn-disabled');
        skipBtn.disabled = false;
    } else {
        document.getElementById('draft-skips').innerText = `見送り不可`;
        skipBtn.classList.add('btn-disabled');
        skipBtn.disabled = true;
    }
}

function startDraft(type) {
    currentDraftType = type;
    requiredCount = type === 'fielder' ? 9 : 15;
    skipsLeft = 5;
    draftPool = type === 'fielder' ? [...fieldersData] : [...pitchersData];
    showScreen('screen-draft');
    nextCandidate();
}

function nextCandidate() {
    const currentCount = currentDraftType === 'fielder' ? myTeam.fielders.length : myTeam.pitchers.length;
    if (currentCount >= requiredCount) {
        if (currentDraftType === 'fielder') startDraft('pitcher');
        else initSetupScreen();
        return;
    }
    if (draftPool.length === 0) {
        alert("候補選手が枯渇しました！");
        return;
    }
    const randIdx = Math.floor(Math.random() * draftPool.length);
    currentCandidate = draftPool[randIdx];
    renderCard(currentCandidate, currentDraftType);
    updateDraftHeader();
}

function takePlayer() {
    if (currentDraftType === 'fielder') myTeam.fielders.push(currentCandidate);
    else myTeam.pitchers.push(currentCandidate);
    draftPool = draftPool.filter(p => p !== currentCandidate);
    nextCandidate();
}

function skipPlayer() {
    if (skipsLeft > 0) {
        skipsLeft--;
        draftPool = draftPool.filter(p => p !== currentCandidate);
        nextCandidate();
    }
}

// ==========================================
// 4. カード描画
// ==========================================
function renderCard(player, type) {
    const name = player["選手名"] || "不明";
    const team = player["チーム"] || "フリー";
    const htmlContainer = document.getElementById('draft-card');
    
    if (type === 'fielder') {
        // ★ 守備力のパース（例："1B:56 / LF:60" から最大値60とポジションリストを取得）
        let defenseMax = 0;
        let posTags = [];
        
        if (typeof player["守備力"] === 'string' && player["守備力"].includes(':')) {
            const parts = player["守備力"].split('/');
            parts.forEach(p => {
                const arr = p.split(':');
                if (arr.length === 2) {
                    const rating = parseInt(arr[1], 10);
                    if (!isNaN(rating)) {
                        if (rating > defenseMax) defenseMax = rating;
                        posTags.push({ pos: arr[0].trim(), val: rating });
                    }
                }
            });
        } else {
            defenseMax = parseInt(player["守備力"], 10) || 0;
        }

        // 守れるポジションのタグHTMLを生成
        let posHtml = '';
        if (posTags.length > 0) {
            posHtml = `
                <div style="border-top: 1px dashed #ccc; margin-top: 15px; padding-top: 15px; display: flex; align-items: center; gap: 10px;">
                    <div style="font-size: 13px; color: #555; font-weight: bold; white-space: nowrap;">守れる所</div>
                    <div class="defense-positions">
                        ${posTags.map(t => {
                            const r = getRankInfo(t.val);
                            return `<span class="pos-tag">${t.pos} <span style="color:${r.color}; font-size:14px;">${r.rank}</span></span>`;
                        }).join('')}
                    </div>
                </div>
            `;
        }

        htmlContainer.innerHTML = `
            <div class="card">
                <div class="p-tag" style="background-color: #00796B;">野手</div>
                <div class="p-name">${name}</div>
                <div class="p-meta">${team}</div>
                <div class="stats-box grid-4">
                    ${createStatCard('ミート', player["ミート"])}
                    ${createStatCard('パワー', player["パワー"])}
                    ${createStatCard('走力', player["走力"])}
                    ${createStatCard('守備力', defenseMax)}
                </div>
                ${posHtml}
            </div>
        `;
    } else {
        let pitchTags = '';
        Object.keys(player).forEach(key => {
            if (key.startsWith('球種_') && player[key]) {
                const pitchName = key.replace('球種_', '');
                pitchTags += `<span class="pitch-tag">${pitchName} ${player[key]}</span>`;
            }
        });

        const pitchHtml = pitchTags ? `
            <div style="border-top: 1px dashed #ccc; margin-top: 15px; padding-top: 15px; display: flex; align-items: center; gap: 10px;">
                <div style="font-size: 13px; color: #555; font-weight: bold; white-space: nowrap;">球種</div>
                <div class="pitch-types">${pitchTags}</div>
            </div>
        ` : '';

        htmlContainer.innerHTML = `
            <div class="card">
                <div class="p-tag" style="background-color: #E53935;">投手</div>
                <div class="p-name">${name}</div>
                <div class="p-meta">${team}</div>
                <div class="stats-box grid-3">
                    ${createStatCard('球威', player["球威"])}
                    ${createStatCard('制球', player["制球"])}
                    ${createStatCard('スタミナ', player["スタミナ"])}
                </div>
                ${pitchHtml}
            </div>
        `;
    }
}

// ==========================================
// 5. オーダー確認 ＆ ペナント実行
// ==========================================
function initSetupScreen() {
    showScreen('screen-setup');
    const fList = myTeam.fielders.map(p => p["選手名"]).join("、");
    document.getElementById('my-fielders-list').innerText = fList;
    const pList = myTeam.pitchers.map(p => p["選手名"]).join("、");
    document.getElementById('my-pitchers-list').innerText = pList;
}

function startPennant() {
    showScreen('screen-result');
    let teams = [
        { name: "あなたのオリジナルチーム", wins: 0, losses: 0 },
        { name: "ライバル球団A", wins: 0, losses: 0 },
        { name: "ライバル球団B", wins: 0, losses: 0 },
        { name: "ライバル球団C", wins: 0, losses: 0 },
        { name: "ライバル球団D", wins: 0, losses: 0 },
        { name: "ライバル球団E", wins: 0, losses: 0 }
    ];

    let baseStrength = 0.45 + (Math.random() * 0.15); 
    for(let i=0; i<143; i++) {
        if(Math.random() < baseStrength) teams[0].wins++;
        else teams[0].losses++;
        for(let j=1; j<6; j++) {
            if(Math.random() < 0.5) teams[j].wins++;
            else teams[j].losses++;
        }
    }

    teams.sort((a, b) => b.wins - a.wins);
    let myRank = teams.findIndex(t => t.name === "あなたのオリジナルチーム") + 1;
    document.getElementById('final-rank').innerHTML = `6チーム中 <span style="color:#c62828; font-size:40px;">${myRank}位</span>`;

    let tableHtml = `<table><tr><th>順位</th><th>チーム名</th><th>勝利</th><th>敗北</th><th>勝率</th></tr>`;
    teams.forEach((t, index) => {
        let wp = (t.wins / (t.wins + t.losses)).toFixed(3).replace(/^0\./, '.');
        let isMyTeam = t.name === "あなたのオリジナルチーム" ? 'style="background-color:#fff9c4; font-weight:bold;"' : '';
        tableHtml += `<tr ${isMyTeam}><td>${index+1}</td><td>${t.name}</td><td>${t.wins}</td><td>${t.losses}</td><td>${wp}</td></tr>`;
    });
    tableHtml += `</table>`;
    document.getElementById('result-standings').innerHTML = tableHtml;
}
