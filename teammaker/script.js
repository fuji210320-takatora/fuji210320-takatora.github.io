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
        const res = await fetch('./baseball_data.xlsx?v=' + new Date().getTime());
        if (!res.ok) throw new Error("Fetch Error");
        
        const buffer = await res.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        
        const rawFielders = XLSX.utils.sheet_to_json(workbook.Sheets['野手能力']) || [];
        const rawPitchers = XLSX.utils.sheet_to_json(workbook.Sheets['投手能力']) || [];
        
        fieldersData = rawFielders.filter(p => p['選手名']);
        pitchersData = rawPitchers.filter(p => p['選手名']);
        
        document.getElementById('screen-loading').classList.remove('active');
        startDraft('fielder');
    } catch (e) {
        console.error("データの読み込みに失敗しました:", e);
        const loadTxt = document.getElementById('loading-text');
        if (loadTxt) loadTxt.innerText = "読み込みエラーです。再読み込みしてください。";
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
    if (draftPool.length === 0) { alert("候補選手が枯渇しました！"); return; }
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
        let defenseMax = 0;
        let posTags = [];
        if (typeof player["守備力"] === 'string' && player["守備力"].includes(':')) {
            player["守備力"].split('/').forEach(p => {
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
        // ★ 投手の球種にもランクと色を適用
        let pitchTags = '';
        Object.keys(player).forEach(key => {
            if (key.startsWith('球種_') && player[key]) {
                const pitchName = key.replace('球種_', '');
                const pitchVal = parseInt(player[key], 10);
                if(!isNaN(pitchVal)) {
                    const r = getRankInfo(pitchVal);
                    pitchTags += `<span class="pitch-tag">${pitchName} <span style="color:${r.color}; font-size:14px;">${r.rank}</span></span>`;
                } else {
                    pitchTags += `<span class="pitch-tag">${pitchName} ${player[key]}</span>`;
                }
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
// 5. 起用設定画面の初期化
// ==========================================
function initSetupScreen() {
    showScreen('screen-setup');
    
    // 野手の設定テーブル作成
    let fHtml = `<table><tr><th style="width:80px;">打順</th><th>選手名</th><th>ミート</th><th>パワー</th></tr>`;
    myTeam.fielders.forEach((p, index) => {
        let options = '';
        for(let i=1; i<=9; i++) {
            let selected = (index + 1 === i) ? 'selected' : '';
            options += `<option value="${i}番" ${selected}>${i}番</option>`;
        }
        fHtml += `<tr>
            <td><select id="f-role-${index}">${options}</select></td>
            <td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${p["ミート"]||'-'}</td><td>${p["パワー"]||'-'}</td>
        </tr>`;
    });
    fHtml += `</table>`;
    document.getElementById('setup-fielders').innerHTML = fHtml;

    // 投手の設定テーブル作成
    let pHtml = `<table><tr><th style="width:100px;">役割</th><th>選手名</th><th>球威</th><th>制球</th></tr>`;
    const pRoles = ['先発1','先発2','先発3','先発4','先発5','先発6','中継ぎ','中継ぎ','中継ぎ','中継ぎ','中継ぎ','中継ぎ','中継ぎ','中継ぎ','抑え'];
    myTeam.pitchers.forEach((p, index) => {
        let options = '';
        ['先発', '中継ぎ', '抑え'].forEach(role => {
            let selected = pRoles[index].includes(role) ? 'selected' : '';
            options += `<option value="${role}" ${selected}>${role}</option>`;
        });
        pHtml += `<tr>
            <td><select id="p-role-${index}">${options}</select></td>
            <td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${p["球威"]||'-'}</td><td>${p["制球"]||'-'}</td>
        </tr>`;
    });
    pHtml += `</table>`;
    document.getElementById('setup-pitchers').innerHTML = pHtml;
}

// ==========================================
// 6. ペナント実行 ＆ 個人成績生成
// ==========================================
function switchTab(tabId) {
    document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    event.target.classList.add('active');
    document.getElementById(tabId).classList.add('active');
}

function startPennant() {
    // ユーザーの設定を選手データに保存
    myTeam.fielders.forEach((p, i) => { p.role = document.getElementById(`f-role-${i}`).value; });
    myTeam.pitchers.forEach((p, i) => { p.role = document.getElementById(`p-role-${i}`).value; });

    showScreen('screen-result');
    
    // --- 順位表の生成 ---
    let teams = [
        { name: "あなたのオリジナルチーム", wins: 0, losses: 0 },
        { name: "ライバル球団A", wins: 0, losses: 0 }, { name: "ライバル球団B", wins: 0, losses: 0 },
        { name: "ライバル球団C", wins: 0, losses: 0 }, { name: "ライバル球団D", wins: 0, losses: 0 },
        { name: "ライバル球団E", wins: 0, losses: 0 }
    ];
    let baseStrength = 0.45 + (Math.random() * 0.15); 
    for(let i=0; i<143; i++) {
        if(Math.random() < baseStrength) teams[0].wins++; else teams[0].losses++;
        for(let j=1; j<6; j++) {
            if(Math.random() < 0.5) teams[j].wins++; else teams[j].losses++;
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
    document.getElementById('tab-standings').innerHTML = tableHtml;

    // --- 個人成績の生成（野手） ---
    let statFHtml = `<table><tr><th>打順</th><th>選手名</th><th>打率</th><th>本塁打</th><th>打点</th><th>OPS</th></tr>`;
    myTeam.fielders.forEach(p => {
        const meet = parseInt(p["ミート"], 10) || 50;
        const power = parseInt(p["パワー"], 10) || 50;
        
        // 能力ベースに乱数でブレ幅を作る
        let avg = 0.200 + ((meet + (Math.random()*20-10)) * 0.0013);
        if(avg > 0.350) avg = 0.340 + Math.random()*0.02; // 上限調整
        let hr = Math.max(0, Math.floor((power - 30 + (Math.random()*20-10)) * 0.45));
        let rbi = hr * 2 + Math.floor(Math.random() * 40) + 15;
        let ops = avg + 0.300 + (hr * 0.012);

        statFHtml += `<tr>
            <td>${p.role}</td><td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${avg.toFixed(3).replace(/^0\./, '.')}</td>
            <td>${hr}</td><td>${rbi}</td>
            <td>${ops.toFixed(3).replace(/^0\./, '.')}</td>
        </tr>`;
    });
    statFHtml += `</table>`;
    document.getElementById('tab-fielders').innerHTML = statFHtml;

    // --- 個人成績の生成（投手） ---
    let statPHtml = `<table><tr><th>役割</th><th>選手名</th><th>防御率</th><th>登板</th><th>勝利</th><th>敗北</th><th>セーブ</th></tr>`;
    myTeam.pitchers.forEach(p => {
        const kyui = parseInt(p["球威"], 10) || 50;
        const seikyu = parseInt(p["制球"], 10) || 50;
        const total = kyui + seikyu + (Math.random()*30-15);
        
        let era = 6.00 - (total * 0.02);
        if(era < 1.00) era = 1.20 + Math.random();
        if(era > 8.00) era = 7.50 + Math.random();

        let games = 0, win = 0, lose = 0, sv = 0;
        if(p.role === '先発') {
            games = Math.floor(Math.random() * 5) + 20; // 20~24登板
            win = Math.floor((10 - era) * 1.5 + Math.random()*4);
            lose = Math.floor(era * 1.2 + Math.random()*4);
        } else if(p.role === '抑え') {
            games = Math.floor(Math.random() * 20) + 40; // 40~59登板
            sv = Math.floor((10 - era) * 4 + Math.random()*10);
            win = Math.floor(Math.random() * 3);
            lose = Math.floor(Math.random() * 4);
        } else { // 中継ぎ
            games = Math.floor(Math.random() * 30) + 30; // 30~59登板
            win = Math.floor(Math.random() * 5);
            lose = Math.floor(Math.random() * 4);
        }

        statPHtml += `<tr>
            <td>${p.role}</td><td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${era.toFixed(2)}</td><td>${games}</td>
            <td>${Math.max(0, win)}</td><td>${Math.max(0, lose)}</td><td>${sv}</td>
        </tr>`;
    });
    statPHtml += `</table>`;
    document.getElementById('tab-pitchers').innerHTML = statPHtml;
}
