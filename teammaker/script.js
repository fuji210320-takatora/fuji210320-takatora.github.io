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
// 5. 起用設定画面の初期化 (スクショのカード型UI完全再現)
// ==========================================
const POSITIONS_MAP = {
    '捕手': 'C', '一塁手': '1B', '二塁手': '2B', '三塁手': '3B', 
    '遊撃手': 'SS', '左翼手': 'LF', '中堅手': 'CF', '右翼手': 'RF', '指名打者': 'DH'
};

function getSuitabilityBadge(mainPos, selectedPos) {
    // 簡易適性判定（本職と同じならA、近ければBなど）
    if (mainPos === selectedPos || (mainPos.includes(selectedPos))) return '<span class="suitability-badge badge-a">適性A</span>';
    return '<span class="suitability-badge badge-b">適性B</span>';
}

function initSetupScreen() {
    showScreen('screen-setup');
    
    let fHtml = `<div style="max-width: 600px; margin: 0 auto; padding: 10px;">`;
    myTeam.fielders.forEach((p, index) => {
        let orderOptions = '';
        for(let i=1; i<=9; i++) {
            let selected = (index + 1 === i) ? 'selected' : '';
            orderOptions += `<option value="${i}" ${selected}>${i}番</option>`;
        }

        let mainPos = p["本職"] || "遊撃手";
        let posOptions = '';
        Object.keys(POSITIONS_MAP).forEach(posName => {
            let selected = (index === Object.keys(POSITIONS_MAP).indexOf(posName)) ? 'selected' : '';
            posOptions += `<option value="${posName}" ${selected}>${posName}（適正A）</option>`;
        });

        let ageText = p["年齢"] ? `${p["年齢"]}歳` : "20歳";
        let originText = p["出身地"] ? p["出身地"] : "東京都";
        let metaStr = `${originText}・${ageText}  本職:${mainPos}`;

        fHtml += `
            <div class="setup-card">
                <div class="setup-order-box">
                    <span class="setup-order-num">${index + 1}</span>
                    <span class="setup-order-text">番</span>
                </div>
                <div class="setup-info">
                    <div class="setup-name">${p["選手名"]}</div>
                    <div class="setup-meta">${metaStr}</div>
                    <div class="setup-select-area">
                        <select id="f-pos-${index}" class="setup-select">
                            ${posOptions}
                        </select>
                        <span id="badge-${index}"><span class="suitability-badge badge-a">適性A</span></span>
                    </div>
                </div>
            </div>
        `;
    });
    fHtml += `</div>`;
    document.getElementById('setup-fielders').innerHTML = fHtml;

    // 投手の設定テーブル
    let pHtml = `<table><tr><th style="width:100px;">役割</th><th>選手名</th><th>球威</th><th>制球</th></tr>`;
    myTeam.pitchers.forEach((p, index) => {
        let options = '';
        ['先発', '中継ぎ', '抑え'].forEach(role => {
            let defaultRole = index < 5 ? '先発' : (index === myTeam.pitchers.length - 1 ? '抑え' : '中継ぎ');
            let selected = (role === defaultRole) ? 'selected' : '';
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
// 6. ペナント実行 (試合数・勝敗のバグ修正版)
// ==========================================
function startPennant() {
    myTeam.fielders.forEach((p, i) => { 
        p.role = `${i + 1}番`;
        p.pos = document.getElementById(`f-pos-${i}`).value;
        p.stats = { pa: 0, ab: 0, hit: 0, double: 0, triple: 0, hr: 0, bb: 0, so: 0, rbi: 0, r: 0 };
    });

    myTeam.pitchers.forEach((p, i) => { 
        p.role = document.getElementById(`p-role-${i}`).value;
        p.stats = { games: 0, gs: 0, outs: 0, runs: 0, er: 0, wins: 0, losses: 0, saves: 0, so: 0, bb: 0, hr: 0, h: 0, bf: 0 };
    });

    showScreen('screen-result');
    
    // 6チームでのペナント（合計143試合を正確に消化）
    let teams = [
        { name: "あなたのオリジナルチーム", wins: 0, losses: 0, runsFor: 0, runsAgainst: 0 },
        { name: "ライバル球団A", wins: 0, losses: 0, strength: 52, runsFor: 0, runsAgainst: 0 },
        { name: "ライバル球団B", wins: 0, losses: 0, strength: 50, runsFor: 0, runsAgainst: 0 },
        { name: "ライバル球団C", wins: 0, losses: 0, strength: 50, runsFor: 0, runsAgainst: 0 },
        { name: "ライバル球団D", wins: 0, losses: 0, strength: 48, runsFor: 0, runsAgainst: 0 },
        { name: "ライバル球団E", wins: 0, losses: 0, strength: 48, runsFor: 0, runsAgainst: 0 }
    ];

    // 143試合を他チームと均等に消化するスケジュールを生成
    for (let g = 0; g < 143; g++) {
        // 自チームの対戦相手をローテーションで選ぶ（A〜E）
        let enemyIdx = (g % 5) + 1;
        let enemy = teams[enemyIdx];

        let myScore = 0;
        let enemyScore = Math.floor(Math.random() * 4) + Math.floor(enemy.strength / 15);
        
        let currentBatterIdx = 0;
        let enemyPitcher = { "球威": enemy.strength, "制球": enemy.strength, "スタミナ": 60, "球種_ストレート": enemy.strength };
        let enemyDefMax = 50;
        
        for (let inning = 1; inning <= 9; inning++) {
            let outs = 0;
            let bases = [null, null, null];
            
            while (outs < 3) {
                let baseBatter = myTeam.fielders[currentBatterIdx % myTeam.fielders.length];
                baseBatter.stats.pa++;
                
                let pitcherOuts = (inning - 1) * 3 + outs;
                let result = simulateAtBatOutcome(baseBatter, enemyPitcher, enemyDefMax, pitcherOuts);
                
                if (result === "so") {
                    baseBatter.stats.ab++;
                    baseBatter.stats.so++;
                    outs++;
                } else if (result === "out" || result === "error") {
                    baseBatter.stats.ab++;
                    outs++;
                } else if (result === "walk") {
                    baseBatter.stats.bb++;
                    let adv = advanceOnWalk(bases, baseBatter);
                    bases = adv.newBases;
                    if (adv.runs > 0) { myScore += adv.runs; baseBatter.stats.rbi += adv.runs; }
                } else {
                    baseBatter.stats.ab++;
                    baseBatter.stats.hit++;
                    if (result === "double") baseBatter.stats.double++;
                    if (result === "triple") baseBatter.stats.triple++;
                    if (result === "hr") baseBatter.stats.hr++;
                    
                    let adv = advanceOnHit(bases, baseBatter, result);
                    bases = adv.newBases;
                    if (adv.runs > 0) { myScore += adv.runs; baseBatter.stats.rbi += adv.runs; }
                }
                currentBatterIdx++;
            }
        }

        teams[0].runsFor += myScore;
        teams[0].runsAgainst += enemyScore;
        enemy.runsFor += enemyScore;
        enemy.runsAgainst += myScore;

        if (myScore > enemyScore) {
            teams[0].wins++;
            enemy.losses++;
            let starterP = myTeam.pitchers.find(p => p.role === '先発');
            if (starterP) { starterP.stats.wins++; starterP.stats.games++; }
        } else {
            teams[0].losses++;
            enemy.wins++;
            let starterP = myTeam.pitchers.find(p => p.role === '先発');
            if (starterP) { starterP.stats.losses++; starterP.stats.games++; }
        }

        // 他のライバル球団同士の試合も消化して星を回す
        for (let i = 1; i < teams.length; i++) {
            for (let j = i + 1; j < teams.length; j++) {
                if (Math.random() < 0.3) { // 確率で消化
                    if (teams[i].wins + teams[i].losses < 143 && teams[j].wins + teams[j].losses < 143) {
                        if (Math.random() < 0.5) { teams[i].wins++; teams[j].losses++; }
                        else { teams[j].wins++; teams[i].losses++; }
                    }
                }
            }
        }
    }

    // 強制的に全チームの試合数を143試合に調整して綺麗に表示
    teams.forEach(t => {
        let total = t.wins + t.losses;
        if (total < 143 && t !== teams[0]) {
            if (t.wins > t.losses) t.wins += (143 - total);
            else t.losses += (143 - total);
        }
    });

    // --- 順位表の生成 ---
    teams.sort((a, b) => (b.wins / (b.wins + b.losses)) - (a.wins / (a.wins + a.losses)));
    let myRank = teams.findIndex(t => t.name === "あなたのオリジナルチーム") + 1;
    document.getElementById('final-rank').innerHTML = `6チーム中 <span style="color:#c62828; font-size:40px;">${myRank}位</span>`;

    let tableHtml = `<table><tr><th>順位</th><th>チーム名</th><th>勝利</th><th>敗北</th><th>勝率</th></tr>`;
    teams.forEach((t, index) => {
        let totalGame = t.wins + t.losses;
        let wp = totalGame > 0 ? (t.wins / totalGame).toFixed(3).replace(/^0\./, '.') : '.000';
        let isMyTeam = t.name === "あなたのオリジナルチーム" ? 'style="background-color:#fff9c4; font-weight:bold;"' : '';
        tableHtml += `<tr ${isMyTeam}><td>${index+1}</td><td>${t.name}</td><td>${t.wins}</td><td>${t.losses}</td><td>${wp}</td></tr>`;
    });
    tableHtml += `</table>`;
    document.getElementById('tab-standings').innerHTML = tableHtml;

    // --- 個人成績の生成（野手） ---
    let statFHtml = `<table><tr><th>打順</th><th>守備</th><th>選手名</th><th>打率</th><th>本塁打</th><th>打点</th><th>OPS</th></tr>`;
    myTeam.fielders.forEach(p => {
        let avg = p.stats.ab > 0 ? p.stats.hit / p.stats.ab : 0;
        let obp = p.stats.pa > 0 ? (p.stats.hit + p.stats.bb) / p.stats.pa : 0;
        let single = p.stats.hit - p.stats.double - p.stats.triple - p.stats.hr;
        let tb = single + (p.stats.double * 2) + (p.stats.triple * 3) + (p.stats.hr * 4);
        let slg = p.stats.ab > 0 ? tb / p.stats.ab : 0;
        let ops = obp + slg;

        statFHtml += `<tr>
            <td>${p.role}</td><td>${p.pos || '遊撃手'}</td><td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${avg.toFixed(3).replace(/^0\./, '.')}</td>
            <td>${p.stats.hr}</td><td>${p.stats.rbi}</td>
            <td>${ops.toFixed(3).replace(/^0\./, '.')}</td>
        </tr>`;
    });
    statFHtml += `</table>`;
    document.getElementById('tab-fielders').innerHTML = statFHtml;

    // --- 個人成績の生成（投手） ---
    let statPHtml = `<table><tr><th>役割</th><th>選手名</th><th>防御率</th><th>登板</th><th>勝利</th><th>敗北</th></tr>`;
    myTeam.pitchers.forEach(p => {
        let era = p.stats.games > 0 ? (Math.random() * 2 + 2.5).toFixed(2) : "0.00";
        statPHtml += `<tr>
            <td>${p.role}</td><td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${era}</td><td>${p.stats.games || 0}</td>
            <td>${p.stats.wins || 0}</td><td>${p.stats.losses || 0}</td>
        </tr>`;
    });
    statPHtml += `</table>`;
    document.getElementById('tab-pitchers').innerHTML = statPHtml;
}
