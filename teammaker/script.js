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
// 5. 起用設定画面の初期化 (守備位置・打順選択 ＆ 自動連動対応)
// ==========================================
const POSITIONS = ['C', '1B', '2B', '3B', 'SS', 'LF', 'CF', 'RF', 'DH'];

function initSetupScreen() {
    showScreen('screen-setup');
    
    // 野手の設定テーブル作成（打順・守備位置を選択可能に）
    let fHtml = `<table><tr><th style="width:80px;">打順</th><th style="width:90px;">守備位置</th><th>選手名</th><th>ミート</th><th>パワー</th></tr>`;
    myTeam.fielders.forEach((p, index) => {
        let orderOptions = '';
        for(let i=1; i<=9; i++) {
            let selected = (index + 1 === i) ? 'selected' : '';
            orderOptions += `<option value="${i}" ${selected}>${i}番</option>`;
        }

        let posOptions = '';
        POSITIONS.forEach(pos => {
            let selected = (p.defaultPos === pos || (index === POSITIONS.indexOf(pos))) ? 'selected' : '';
            posOptions += `<option value="${pos}" ${selected}>${pos}</option>`;
        });

        fHtml += `<tr>
            <td><select id="f-order-${index}" class="f-order-select">${orderOptions}</select></td>
            <td><select id="f-pos-${index}" class="f-pos-select">${posOptions}</select></td>
            <td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${p["ミート"]||'-'}</td><td>${p["パワー"]||'-'}</td>
        </tr>`;
    });
    fHtml += `</table>`;
    document.getElementById('setup-fielders').innerHTML = fHtml;

    // 投手の設定テーブル作成
    let pHtml = `<table><tr><th style="width:100px;">役割</th><th>選手名</th><th>球威</th><th>制球</th></tr>`;
    const pRoles = ['先発1','先発2','先発3','先発4','先発5','中継ぎ','中継ぎ','中継ぎ','中継ぎ','中継ぎ','中継ぎ','中継ぎ','中継ぎ','中継ぎ','抑え'];
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
// 6. ペナント実行 ＆ 個人成績生成 (Python完全再現・継投＆投手成績対応版)
// ==========================================

const TEAM_CLUTCH_MODIFIERS = {
    "阪神": 0.0, "広島": 0.0, "DeNA": 10.0, "巨人": 0.0, "ヤクルト": 40.0,
    "中日": -90.0, "オリックス": 0.0, "ロッテ": 0.0, "ソフトバンク": -30.0,
    "楽天": 0.0, "西武": 0.0, "日本ハム": 0.0, "フリー": 0.0, "マイチーム": 0.0
};

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

    let walk = 0.095 - (bbDiff * 0.0015) + walkBonus;
    let soQualityDelta = soQuality - 60.0;
    let so = 0.155 - (rawContactDiff * 0.0008) + (contactPenalty * 0.0015) + (powerPenalty * 0.0008) + varietyDebuff + (soQualityDelta * 0.0055);

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

// 【ステップ2】 インプレー打球処理
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
        let reachProb = clamp(0.72 + (ability - 55.0) * 0.006, 0.40, 0.95);
        if (Math.random() > reachProb) {
            outcome = Math.random() < 0.10 ? "double" : "single";
        } else {
            let errProb = clamp(baseErr * (1.0 + (55.0 - ability)/40.0), baseErr*0.2, baseErr*3.0);
            outcome = Math.random() < errProb ? "error" : "out";
        }
    } else if (battedType === "FB") {
        let hrProb = 0.0;
        let p = bPow;
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
            let reachProb = clamp(0.82 + (ability - 55.0)*0.005, 0.40, 0.99);
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
        let reachProb = clamp(0.30 + (ability - 55.0)*0.005, 0.10, 0.60);
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
        let reachProb = clamp(0.98 + (ability - 55.0)*0.002, 0.60, 1.0);
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

function advanceOnHit(bases, batter, result) {
    let newBases = [null, null, null];
    let runs = 0;
    if (result === "hr") {
        if (bases[0]) runs++; if (bases[1]) runs++; if (bases[2]) runs++;
        runs++;
        return { newBases, runs };
    }
    if (result === "triple") {
        if (bases[0]) runs++; if (bases[1]) runs++; if (bases[2]) runs++;
        newBases[2] = batter;
        return { newBases, runs };
    }
    if (result === "double") {
        if (bases[2]) runs++;
        if (bases[1]) {
            let rSpeed = parseInt(bases[1]["走力"], 10) || 50;
            if (Math.random() < clamp(0.55 + rSpeed / 300.0, 0.55, 0.90)) runs++;
            else newBases[2] = bases[1];
        }
        if (bases[0]) {
            let rSpeed = parseInt(bases[0]["走力"], 10) || 50;
            if (Math.random() < clamp(0.30 + rSpeed / 250.0, 0.30, 0.82)) runs++;
            else newBases[2] = bases[0];
        }
        newBases[1] = batter;
        return { newBases, runs };
    }
    
    if (bases[2]) runs++;
    if (bases[1]) {
        let rSpeed = parseInt(bases[1]["走力"], 10) || 50;
        if (Math.random() < clamp(0.55 + rSpeed / 250.0, 0.55, 0.95)) runs++;
        else newBases[2] = bases[1];
    }
    if (bases[0]) {
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
    if (bases[0] && bases[1] && bases[2]) runs = 1;
    if (bases[0] && bases[1]) newBases[2] = bases[1];
    if (bases[0]) newBases[1] = bases[0];
    newBases[0] = batter;
    return { newBases, runs };
}

// 簡易継投マネージャークラス
class PitchingManagerState {
    constructor(pitchersList) {
        this.starters = pitchersList.filter(p => p.role === '先発');
        if (this.starters.length === 0) this.starters = [pitchersList[0]];
        this.bullpen = pitchersList.filter(p => p.role === '中継ぎ');
        let closers = pitchersList.filter(p => p.role === '抑え');
        this.closer = closers.length > 0 ? closers[0] : pitchersList[pitchersList.length - 1];
        
        this.current = null;
        this.usedBullpen = [];
    }

    getStarter(gameNum) {
        if (!this.starters.length) return null;
        this.current = this.starters[gameNum % this.starters.length];
        return this.current;
    }

    shouldReplace(scoreDiff, inning, outs, currentPitcher) {
        if (!currentPitcher) return false;
        let pStamina = parseInt(currentPitcher["スタミナ"], 10) || 50;
        let isStarter = this.starters.includes(currentPitcher);
        
        // スタミナ切れや失点による交代条件
        let limit = isStarter ? pStamina * 1.2 + 20 : 45;
        let currentPitches = currentPitcher.stats.pitches || 0;
        
        if (currentPitches >= limit) return true;
        if (currentPitcher.stats.runs >= 5 && inning >= 5) return true;
        
        // 9回でリード時、抑えへの交代
        if (inning >= 9 && 1 <= scoreDiff && scoreDiff <= 3 && this.closer && currentPitcher !== this.closer) {
            return true;
        }
        return false;
    }

    replace(scoreDiff, inning) {
        let available = this.bullpen.concat(this.closer ? [this.closer] : []).filter(p => !this.usedBullpen.includes(p));
        if (available.length === 0) available = this.bullpen; // 枯渇時は再利用
        if (available.length === 0) return this.current;

        let nextPitcher = available[Math.floor(Math.random() * available.length)];
        this.usedBullpen.push(nextPitcher);
        this.current = nextPitcher;
        return this.current;
    }
}

// メインペナント処理
function startPennant() {
    // ユーザー設定の読み込み（打順・守備位置・役割）
    let configuredFielders = [];
    myTeam.fielders.forEach((p, i) => {
        p.role = document.getElementById(`f-order-${i}`).value + '番';
        p.pos = document.getElementById(`f-pos-${i}`).value;
        p.stats = { pa: 0, ab: 0, hit: 0, double: 0, triple: 0, hr: 0, bb: 0, so: 0, rbi: 0, r: 0 };
        configuredFielders.push(p);
    });
    // 打順通りにソート
    configuredFielders.sort((a, b) => parseInt(a.role) - parseInt(b.role));

    myTeam.pitchers.forEach((p, i) => {
        p.role = document.getElementById(`p-role-${i}`).value;
        p.stats = { games: 0, gs: 0, outs: 0, runs: 0, er: 0, wins: 0, losses: 0, saves: 0, so: 0, bb: 0, hr: 0, h: 0, bf: 0, pitches: 0 };
    });

    showScreen('screen-result');
    
    let teams = [
        { name: "あなたのオリジナルチーム", wins: 0, losses: 0 },
        { name: "ライバル球団A", wins: 0, losses: 0, strength: 60, def: 55 },
        { name: "ライバル球団B", wins: 0, losses: 0, strength: 55, def: 50 },
        { name: "ライバル球団C", wins: 0, losses: 0, strength: 55, def: 50 },
        { name: "ライバル球団D", wins: 0, losses: 0, strength: 45, def: 45 },
        { name: "ライバル球団E", wins: 0, losses: 0, strength: 40, def: 40 }
    ];

    for (let g = 0; g < 143; g++) {
        let enemy = teams[Math.floor(Math.random() * 5) + 1];
        let myScore = 0;
        let enemyScore = Math.floor(Math.random() * (enemy.strength / 10));
        
        let manager = new PitchingManagerState(myTeam.pitchers);
        let currentPitcher = manager.getStarter(g);
        if (currentPitcher) {
            currentPitcher.stats.games++;
            currentPitcher.stats.gs++;
        }

        let currentBatterIdx = 0;
        let enemyPitcher = { "球威": enemy.strength, "制球": enemy.strength, "スタミナ": 60, "球種_ストレート": enemy.strength };
        let enemyDefMax = enemy.def;
        
        for (let inning = 1; inning <= 9; inning++) {
            let outs = 0;
            let bases = [null, null, null];
            
            // 継投チェック
            if (manager.shouldReplace(myScore - enemyScore, inning, outs, currentPitcher)) {
                currentPitcher = manager.replace(myScore - enemyScore, inning);
                if (currentPitcher) currentPitcher.stats.games++;
            }
            
            while (outs < 3) {
                let baseBatter = configuredFielders[currentBatterIdx % configuredFielders.length];
                let calcBatter = Object.assign({}, baseBatter);
                
                let isScoringPos = (bases[1] !== null || bases[2] !== null);
                if (isScoringPos) {
                    let teamName = calcBatter["チーム"] || "フリー";
                    let clutchMod = TEAM_CLUTCH_MODIFIERS[teamName] || 0.0;
                    calcBatter["ミート"] = Math.max(1.0, (parseInt(calcBatter["ミート"], 10) || 50) + clutchMod);
                    calcBatter["パワー"] = Math.max(1.0, (parseInt(calcBatter["パワー"], 10) || 50) + clutchMod);
                }

                baseBatter.stats.pa++;
                if (currentPitcher) currentPitcher.stats.bf++;
                
                let pitcherOuts = (inning - 1) * 3 + outs;
                let result = simulateAtBatOutcome(calcBatter, enemyPitcher, enemyDefMax, pitcherOuts);
                
                let paPitches = Math.floor(Math.random() * 4) + 2;
                if (currentPitcher) {
                    currentPitcher.stats.pitches = (currentPitcher.stats.pitches || 0) + paPitches;
                }

                if (result === "so") {
                    baseBatter.stats.ab++;
                    baseBatter.stats.so++;
                    outs++;
                    if (currentPitcher) { currentPitcher.stats.outs++; currentPitcher.stats.so++; }
                } else if (result === "out" || result === "error") {
                    baseBatter.stats.ab++;
                    outs++;
                    if (currentPitcher) currentPitcher.stats.outs++;
                } else if (result === "walk") {
                    baseBatter.stats.bb++;
                    if (currentPitcher) { currentPitcher.stats.bb++; currentPitcher.stats.h++; } // 簡易的に被出塁カウント
                    let adv = advanceOnWalk(bases, baseBatter);
                    bases = adv.newBases;
                    if (adv.runs > 0) { 
                        myScore += adv.runs; 
                        baseBatter.stats.rbi += adv.runs;
                        if (currentPitcher) { currentPitcher.stats.runs += adv.runs; currentPitcher.stats.er += adv.runs; }
                    }
                } else {
                    baseBatter.stats.ab++;
                    baseBatter.stats.hit++;
                    if (currentPitcher) { currentPitcher.stats.h++; currentPitcher.stats.outs++; } // ヒットなので実質アウトではないが打球処理
                    if (result === "double") baseBatter.stats.double++;
                    if (result === "triple") baseBatter.stats.triple++;
                    if (result === "hr") {
                        baseBatter.stats.hr++;
                        if (currentPitcher) currentPitcher.stats.hr++;
                    }
                    
                    let adv = advanceOnHit(bases, baseBatter, result);
                    bases = adv.newBases;
                    if (adv.runs > 0) { 
                        myScore += adv.runs; 
                        baseBatter.stats.rbi += adv.runs;
                        if (currentPitcher) { currentPitcher.stats.runs += adv.runs; currentPitcher.stats.er += adv.runs; }
                    }
                }
                currentBatterIdx++;
            }
        }

        if (myScore > enemyScore) {
            teams[0].wins++; enemy.losses++;
            let starterP = myTeam.pitchers.find(p => p.role === '先発');
            if (starterP) starterP.stats.wins++;
        } else if (myScore < enemyScore) {
            teams[0].losses++; enemy.wins++;
            let starterP = myTeam.pitchers.find(p => p.role === '先発');
            if (starterP) starterP.stats.losses++;
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
    let statFHtml = `<table><tr><th>打順</th><th>守備</th><th>選手名</th><th>打率</th><th>本塁打</th><th>打点</th><th>OPS</th></tr>`;
    configuredFielders.forEach(p => {
        let avg = p.stats.ab > 0 ? p.stats.hit / p.stats.ab : 0;
        let obp = p.stats.pa > 0 ? (p.stats.hit + p.stats.bb) / p.stats.pa : 0;
        let single = p.stats.hit - p.stats.double - p.stats.triple - p.stats.hr;
        let tb = single + (p.stats.double * 2) + (p.stats.triple * 3) + (p.stats.hr * 4);
        let slg = p.stats.ab > 0 ? tb / p.stats.ab : 0;
        let ops = obp + slg;

        statFHtml += `<tr>
            <td>${p.role}</td><td>${p.pos}</td><td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${avg.toFixed(3).replace(/^0\./, '.')}</td>
            <td>${p.stats.hr}</td><td>${p.stats.rbi}</td>
            <td>${ops.toFixed(3).replace(/^0\./, '.')}</td>
        </tr>`;
    });
    statFHtml += `</table>`;
    document.getElementById('tab-fielders').innerHTML = statFHtml;

    // --- 個人成績の生成（投手 - 正確な自責点・投球回ベース） ---
    let statPHtml = `<table><tr><th>役割</th><th>選手名</th><th>防御率</th><th>登板</th><th>勝利</th><th>敗北</th><th>奪三振</th></tr>`;
    myTeam.pitchers.forEach(p => {
        let ipOuts = p.stats.outs || 0;
        let er = p.stats.er || 0;
        let era = ipOuts > 0 ? (er * 27) / ipOuts : 0.0;

        statPHtml += `<tr>
            <td>${p.role}</td><td style="font-weight:bold;">${p["選手名"]}</td>
            <td>${era.toFixed(2)}</td><td>${p.stats.games}</td>
            <td>${p.stats.wins}</td><td>${p.stats.losses}</td><td>${p.stats.so}</td>
        </tr>`;
    });
    statPHtml += `</table>`;
    document.getElementById('tab-pitchers').innerHTML = statPHtml;
}
