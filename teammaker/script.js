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
// 6. ペナント実行 ＆ 個人成績生成 (Python完全再現版)
// ==========================================

// ▼ チームごとの「得点圏（チャンス）」裏ステータス設定
const TEAM_CLUTCH_MODIFIERS = {
    "阪神": 0.0, "広島": 0.0, "DeNA": 10.0, "巨人": 0.0, "ヤクルト": 40.0,
    "中日": -90.0, "オリックス": 0.0, "ロッテ": 0.0, "ソフトバンク": -30.0,
    "楽天": 0.0, "西武": 0.0, "日本ハム": 0.0, "フリー": 0.0, "マイチーム": 0.0
};

// ヘルパー関数
function clamp(x, lo, hi) { return Math.max(lo, Math.min(hi, x)); }

function fatigueFactor(stamina, gameOuts) {
    let innings = gameOuts / 3.0;
    if (innings <= 4) return 1.0;
    let excess = innings - 4;
    let penalty = excess * Math.max(0.0, (80.0 - stamina)) * 0.0015;
    return clamp(1.0 - penalty, 0.82, 1.0);
}

// 【ステップ1】 三振・四球・インプレーの3択
function atBatProbabilities(batter, pitcher, gameOuts) {
    let bCon = parseInt(batter["ミート"], 10) || 50;
    let bPow = parseInt(batter["パワー"], 10) || 50;
    let pPow = parseInt(pitcher["球威"], 10) || 50;
    let pCtrl = parseInt(pitcher["制球"], 10) || 50;
    let pStamina = parseInt(pitcher["スタミナ"], 10) || 50;

    // 球種数と平均クオリティの概算（データ構造に合わせる）
    let pitchQuality = 50.0;
    let pitchVariety = 0;
    Object.keys(pitcher).forEach(k => {
        if (k.startsWith('球種_') && parseInt(pitcher[k], 10)) {
            pitchVariety++;
            pitchQuality = (pitchQuality + parseInt(pitcher[k], 10)) / 2;
        }
    });
    if (pitchVariety === 0) pitchVariety = 1;

    let soQuality = (pitchQuality * 0.4) + (pPow * 0.4) + (pCtrl * 0.2) + 5.0;
    let bbQuality = (pCtrl * 0.8) + (pitchQuality * 0.2) + 5.0;

    let fatigue = fatigueFactor(pStamina, gameOuts);
    let bbDiff = bbQuality - 50.0;
    let rawContactDiff = bCon - soQuality;

    let contactPenalty = 0.0;
    if (bCon < 45.0) {
        let diff = 45.0 - bCon;
        contactPenalty = (diff * 0.4) + (Math.pow(diff, 2) * 0.01);
    }
    let powerPenalty = 0.0;
    if (bPow < 60.0) {
        let diff = 60.0 - bPow;
        powerPenalty = (diff * 0.4) + (Math.pow(diff, 2) * 0.01);
    }

    let varietyDebuff = Math.max(0, pitchVariety - 2) * 0.0015;

    let walkBonus = 0.0;
    if (bPow > 50.0) walkBonus += (bPow - 50.0) * 0.0012;
    if (bCon > 50.0) walkBonus += (bCon - 50.0) * 0.0006;
    if (bPow >= 80.0) walkBonus += (bPow - 80.0) * 0.0015;

    let baseWalk = 0.095;
    let baseSo = 0.155;

    let walk = baseWalk - (bbDiff * 0.0015) + walkBonus;
    let soQualityDelta = soQuality - 60.0;
    let so = baseSo - (rawContactDiff * 0.0008) + (contactPenalty * 0.0015) + (powerPenalty * 0.0008) + varietyDebuff + (soQualityDelta * 0.0055);

    if (fatigue < 1.0) {
        walk += (1.0 - fatigue) * 0.02;
        so -= (1.0 - fatigue) * 0.03;
    }

    walk = clamp(walk, 0.010, 0.20);
    so = clamp(so, 0.05, 0.45);
    let inPlay = Math.max(0.20, 1.0 - (walk + so));

    let totalStep1 = walk + so + inPlay;
    return [
        { result: "walk", prob: walk / totalStep1 },
        { result: "so", prob: so / totalStep1 },
        { result: "in_play", prob: inPlay / totalStep1 }
    ];
}

// 【ステップ2】 インプレー打球の処理
function resolveStatcastInPlay(batter, pitcher, ability) {
    let bCon = parseInt(batter["ミート"], 10) || 50;
    let bPow = parseInt(batter["パワー"], 10) || 50;
    let pPow = parseInt(pitcher["球威"], 10) || 50;
    let pCtrl = parseInt(pitcher["制球"], 10) || 50;

    let gbProb = 0.45, fbProb = 0.25, ldProb = 0.20, puProb = 0.10;

    let bPowDiff = bPow - 50.0;
    if (bPowDiff > 0) { fbProb += bPowDiff * 0.0015; gbProb -= bPowDiff * 0.0015; }
    
    let bConDiff = bCon - 50.0;
    if (bConDiff > 0) { ldProb += bConDiff * 0.0025; puProb -= bConDiff * 0.0025; }

    let ctrlDiff = pCtrl - 50.0;
    gbProb += ctrlDiff * 0.0020; ldProb -= ctrlDiff * 0.0020;
    let powDiff = pPow - 50.0;
    puProb += powDiff * 0.0020; fbProb -= powDiff * 0.0020;

    gbProb = Math.max(0.05, gbProb);
    fbProb = Math.max(0.05, fbProb);
    ldProb = Math.max(0.05, ldProb);
    puProb = Math.max(0.01, puProb);

    let tot = gbProb + fbProb + ldProb + puProb;
    let r = Math.random() * tot;
    let battedType = "GB";
    if (r < gbProb) battedType = "GB";
    else if (r < gbProb + fbProb) battedType = "FB";
    else if (r < gbProb + fbProb + ldProb) battedType = "LD";
    else battedType = "PU";

    let baseErr = 0.02;
    let outcome = "out";

    if (battedType === "GB") {
        let baseReach = 0.72; // ▼ 打率アップ修正③
        let reachProb = clamp(baseReach + (ability - 55.0) * 0.006, 0.40, 0.95);
        if (Math.random() > reachProb) {
            outcome = Math.random() < 0.10 ? "double" : "single";
        } else {
            let errProb = clamp(baseErr * (1.0 + (55.0 - ability)/40.0), baseErr*0.2, baseErr*3.0);
            outcome = Math.random() < errProb ? "error" : "out";
        }
    } else if (battedType === "FB") {
        let hrProb = 0.0;
        let p = bPow;
        // ▼ パワーによる専用傾斜（完全再現）
        if (p < 50.0) hrProb = Math.max(0.0, (p - 30.0) * 0.002);
        else if (p < 60.0) hrProb = 0.030 + (p - 50.0) * 0.0040;
        else if (p < 65.0) hrProb = 0.070 + (p - 60.0) * 0.0050;
        else if (p < 70.0) hrProb = 0.095 + (p - 65.0) * 0.0060;
        else if (p < 75.0) hrProb = 0.125 + (p - 70.0) * 0.0080;
        else if (p < 77.0) hrProb = 0.165 + (p - 75.0) * 0.0125;
        else if (p < 80.0) hrProb = 0.190 + (p - 77.0) * 0.0100;
        else if (p < 85.0) hrProb = 0.220 + (p - 80.0) * 0.0080;
        else if (p < 90.0) hrProb = 0.260 + (p - 85.0) * 0.0080;
        else hrProb = 0.300 + (p - 90.0) * 0.0050;

        hrProb = clamp(hrProb, 0.0, 0.40);
        
        if (Math.random() < hrProb) outcome = "hr";
        else {
            let baseReach = 0.82; // ▼ 打率アップ修正④
            let reachProb = clamp(baseReach + (ability - 55.0)*0.005, 0.40, 0.99);
            if (Math.random() > reachProb) {
                let hR = Math.random();
                if (hR < 0.65) outcome = "single";
                else if (hR < 0.95) outcome = "double";
                else outcome = "triple";
            } else {
                let errProb = clamp(baseErr * (1.0 + (55.0 - ability)/40.0), baseErr*0.2, baseErr*3.0);
                outcome = Math.random() < errProb ? "error" : "out";
            }
        }
    } else if (battedType === "LD") {
        let baseReach = 0.30;
        let reachProb = clamp(baseReach + (ability - 55.0)*0.005, 0.10, 0.60);
        if (Math.random() > reachProb) {
            let hR = Math.random();
            if (hR < 0.70) outcome = "single";
            else if (hR < 0.95) outcome = "double";
            else outcome = "triple";
        } else {
            let errProb = clamp(baseErr * (1.0 + (55.0 - ability)/40.0), baseErr*0.2, baseErr*3.0);
            outcome = Math.random() < errProb ? "error" : "out";
        }
    } else if (battedType === "PU") {
        let baseReach = 0.98;
        let reachProb = clamp(baseReach + (ability - 55.0)*0.002, 0.60, 1.0);
        if (Math.random() > reachProb) outcome = "single";
        else {
            let errProb = clamp(baseErr * (1.0 + (55.0 - ability)/40.0), baseErr*0.2, baseErr*3.0);
            outcome = Math.random() < errProb ? "error" : "out";
        }
    }
    return outcome;
}

function simulateAtBatOutcome(batter, pitcher, defMax, gameOuts) {
    let probs = atBatProbabilities(batter, pitcher, gameOuts);
    let r1 = Math.random();
    let cumulative = 0.0;
    let step1Result = "in_play";
    for (let p of probs) {
        cumulative += p.prob;
        if (r1 <= cumulative) {
            step1Result = p.result;
            break;
        }
    }
    if (step1Result === "so" || step1Result === "walk") return step1Result;
    return resolveStatcastInPlay(batter, pitcher, defMax);
}

// 走者進行処理（Python完全再現）
function advanceOnHit(bases, batter, result) {
    let newBases = [null, null, null];
    let runs = 0;
    if (result === "hr") {
        if (bases[0] !== null) runs++;
        if (bases[1] !== null) runs++;
        if (bases[2] !== null) runs++;
        runs++; // 打者
        return { newBases, runs };
    }
    if (result === "triple") {
        if (bases[0] !== null) runs++;
        if (bases[1] !== null) runs++;
        if (bases[2] !== null) runs++;
        newBases[2] = batter;
        return { newBases, runs };
    }
    if (result === "double") {
        if (bases[2] !== null) runs++;
        if (bases[1] !== null) {
            let rSpeed = parseInt(bases[1]["走力"], 10) || 50;
            if (Math.random() < clamp(0.55 + rSpeed / 300.0, 0.55, 0.90)) runs++;
            else newBases[2] = bases[1];
        }
        if (bases[0] !== null) {
            let rSpeed = parseInt(bases[0]["走力"], 10) || 50;
            if (Math.random() < clamp(0.30 + rSpeed / 250.0, 0.30, 0.82)) runs++;
            else newBases[2] = bases[0];
        }
        newBases[1] = batter;
        return { newBases, runs };
    }
    
    // single / error
    if (bases[2] !== null) runs++;
    if (bases[1] !== null) {
        let rSpeed = parseInt(bases[1]["走力"], 10) || 50;
        if (Math.random() < clamp(0.55 + rSpeed / 250.0, 0.55, 0.95)) runs++;
        else newBases[2] = bases[1];
    }
    if (bases[0] !== null) {
        let rSpeed = parseInt(bases[0]["走力"], 10) || 50;
        if (Math.random() < clamp(0.25 + rSpeed / 300.0, 0.10, 0.70)) newBases[2] = bases[0];
        else newBases[1] = bases[0];
    }
    newBases[0] = batter;
    return { newBases, runs };
}

function advanceOnWalk(bases, batter) {
    let newBases = [...bases];
    let runs = 0;
    if (bases[0] !== null && bases[1] !== null && bases[2] !== null) runs = 1;
    if (bases[0] !== null && bases[1] !== null) newBases[2] = bases[1];
    if (bases[0] !== null) newBases[1] = bases[0];
    newBases[0] = batter;
    return { newBases, runs };
}

function switchTab(tabId) {
    document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    event.target.classList.add('active');
    document.getElementById(tabId).classList.add('active');
}

// -----------------------------------------------------------
// ペナント本体
// -----------------------------------------------------------
function startPennant() {
    myTeam.fielders.forEach((p, i) => { 
        p.role = document.getElementById(`f-role-${i}`).value;
        p.stats = { pa: 0, ab: 0, hit: 0, double: 0, triple: 0, hr: 0, bb: 0, so: 0, rbi: 0 };
    });
    myTeam.pitchers.forEach((p, i) => { 
        p.role = document.getElementById(`p-role-${i}`).value;
        p.stats = { games: 0, outs: 0, runs: 0, er: 0, wins: 0, losses: 0, saves: 0, so: 0, bb: 0, hr: 0 };
    });

    showScreen('screen-result');
    
    let teams = [
        { name: "あなたのオリジナルチーム", wins: 0, losses: 0 },
        { name: "ライバル球団A", wins: 0, losses: 0, strength: 60, def: 55, nameStr: "阪神" },
        { name: "ライバル球団B", wins: 0, losses: 0, strength: 55, def: 50, nameStr: "DeNA" },
        { name: "ライバル球団C", wins: 0, losses: 0, strength: 55, def: 50, nameStr: "巨人" },
        { name: "ライバル球団D", wins: 0, losses: 0, strength: 45, def: 45, nameStr: "中日" },
        { name: "ライバル球団E", wins: 0, losses: 0, strength: 40, def: 40, nameStr: "ヤクルト" }
    ];

    for (let g = 0; g < 143; g++) {
        let enemy = teams[Math.floor(Math.random() * 5) + 1];
        let myScore = 0;
        let enemyScore = Math.floor(Math.random() * (enemy.strength / 10)); // 簡易的な敵の得点
        
        let currentBatterIdx = 0;
        let enemyPitcher = { "球威": enemy.strength, "制球": enemy.strength, "スタミナ": 60, "球種_ストレート": enemy.strength };
        let enemyDefMax = enemy.def;
        
        for (let inning = 1; inning <= 9; inning++) {
            let outs = 0;
            let bases = [null, null, null];
            
            while (outs < 3) {
                let baseBatter = myTeam.fielders[currentBatterIdx % myTeam.fielders.length];
                let calcBatter = Object.assign({}, baseBatter);
                
                // ▼ 追加：得点圏（チャンス）での裏ステータス適用
                let isScoringPos = (bases[1] !== null || bases[2] !== null);
                if (isScoringPos) {
                    let teamName = calcBatter["チーム"] || "フリー";
                    let clutchMod = TEAM_CLUTCH_MODIFIERS[teamName] || 0.0;
                    calcBatter["ミート"] = Math.max(1.0, (parseInt(calcBatter["ミート"], 10) || 50) + clutchMod);
                    calcBatter["パワー"] = Math.max(1.0, (parseInt(calcBatter["パワー"], 10) || 50) + clutchMod);
                }

                baseBatter.stats.pa++;
                
                // 仮想の現在アウト数（スタミナ減少用）
                let pitcherOuts = (inning - 1) * 3 + outs;
                
                let result = simulateAtBatOutcome(calcBatter, enemyPitcher, enemyDefMax, pitcherOuts);
                let currentPitcher = myTeam.pitchers[0]; // TODO: 本格的な投手継投ロジック拡張
                if (currentPitcher) currentPitcher.stats.outs++;

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

        if (myScore > enemyScore) {
            teams[0].wins++; enemy.losses++;
            if (myTeam.pitchers[0]) myTeam.pitchers[0].stats.wins++;
        } else if (myScore < enemyScore) {
            teams[0].losses++; enemy.wins++;
            if (myTeam.pitchers[0]) myTeam.pitchers[0].stats.losses++;
        }
    }

    // --- 順位表の生成 ---
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
        let avg = p.stats.ab > 0 ? p.stats.hit / p.stats.ab : 0;
        let obp = p.stats.pa > 0 ? (p.stats.hit + p.stats.bb) / p.stats.pa : 0;
        let single = p.stats.hit - p.stats.double - p.stats.triple - p.stats.hr;
        let tb = single + (p.stats.double * 2) + (p.stats.triple * 3) + (p.stats.hr * 4);
        let slg = p.stats.ab > 0 ? tb / p.stats.ab : 0;
        let ops = obp + slg;

        statFHtml += `<tr>
            <td>${p.role}</td><td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${avg.toFixed(3).replace(/^0\./, '.')}</td>
            <td>${p.stats.hr}</td><td>${p.stats.rbi}</td>
            <td>${ops.toFixed(3).replace(/^0\./, '.')}</td>
        </tr>`;
    });
    statFHtml += `</table>`;
    document.getElementById('tab-fielders').innerHTML = statFHtml;

    // --- 個人成績の生成（投手 - イニングベースの簡易版） ---
    let statPHtml = `<table><tr><th>役割</th><th>選手名</th><th>防御率</th><th>登板</th><th>勝利</th><th>敗北</th><th>セーブ</th></tr>`;
    myTeam.pitchers.forEach(p => {
        let kyui = parseInt(p["球威"], 10) || 50;
        let seikyu = parseInt(p["制球"], 10) || 50;
        let eraBase = 5.50 - ((kyui + seikyu) * 0.02); 
        let era = Math.max(1.50, eraBase + (Math.random() * 2 - 1));

        statPHtml += `<tr>
            <td>${p.role}</td><td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${era.toFixed(2)}</td>
            <td>${p.stats.wins + p.stats.losses + Math.floor(Math.random()*10)}</td>
            <td>${p.stats.wins}</td><td>${p.stats.losses}</td><td>${p.stats.saves}</td>
        </tr>`;
    });
    statPHtml += `</table>`;
    document.getElementById('tab-pitchers').innerHTML = statPHtml;
}
