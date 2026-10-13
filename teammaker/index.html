<!DOCTYPE html>
<html lang="ja">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>ドラフトシミュレーター</title>
    <!-- Excel読み込みライブラリ -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body {
            font-family: "Helvetica Neue", Arial, "Hiragino Kaku Gothic ProN", "Hiragino Sans", Meiryo, sans-serif;
            background-color: #f6f5f0;
            color: #333;
            margin: 0;
            padding: 20px;
        }

        .screen { display: none; }
        .screen.active { display: block; }
        
        .container {
            max-width: 650px;
            margin: 0 auto;
            background: #fff;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }

        #loading-text { font-size: 18px; font-weight: bold; color: #555; text-align: center; padding: 40px; }
        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 2px solid #eee; padding-bottom: 10px;}
        .draft-info { font-size: 18px; font-weight: bold; }
        .btn-group { display: flex; gap: 10px; margin-top: 20px; }
        
        button {
            flex: 1; padding: 15px; font-size: 16px; font-weight: bold;
            border: none; border-radius: 5px; cursor: pointer; color: white; transition: 0.2s;
        }
        .btn-take { background-color: #1976D2; }
        .btn-take:hover { background-color: #1565C0; }
        .btn-skip { background-color: #757575; }
        .btn-skip:hover { background-color: #616161; }
        .btn-disabled { background-color: #e0e0e0 !important; color: #9e9e9e !important; cursor: not-allowed; }
        .btn-primary { background-color: #2E7D32; margin-top: 20px; width: 100%; }

        /* ドラフト用カードデザイン */
        .card { border: 1px solid #ddd; border-radius: 8px; padding: 20px; background: #fafafa; }
        .p-tag { display: inline-block; padding: 4px 12px; border-radius: 20px; background: #00796B; color: white; font-size: 12px; font-weight: bold; margin-bottom: 10px; }
        .p-name { font-size: 26px; font-weight: bold; margin-bottom: 5px; letter-spacing: 1px;}
        .p-meta { font-size: 14px; color: #666; margin-bottom: 15px; }

        .stats-box { display: grid; gap: 8px; margin-top: 15px; margin-bottom: 15px; }
        .grid-4 { grid-template-columns: repeat(4, 1fr); }
        .grid-3 { grid-template-columns: repeat(3, 1fr); }
        .stat-card { background-color: #f2efe9; border: 1px solid #dcd9d2; border-radius: 6px; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 10px 0; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
        .stat-label { font-size: 11px; color: #555; font-weight: bold; margin-bottom: 4px; }
        .stat-rank { font-size: 28px; font-weight: 900; line-height: 1; margin-bottom: 4px; font-family: "Arial Black", "Impact", sans-serif; -webkit-text-stroke: 0.5px rgba(0,0,0,0.1); }
        .stat-val { font-size: 13px; color: #666; font-weight: bold; text-align: center; }

        .pitch-types, .defense-positions { display: flex; flex-wrap: wrap; gap: 6px; }
        .pos-tag, .pitch-tag { background-color: #e8f5e9; color: #1b5e20; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; border: 1px solid #c8e6c9; }

        /* ======== 起用画面カードデザイン (Screenshot_20261013-133528.png 再現) ======== */
        .setup-card { display: flex; background: #fff; border: 1px solid #e0e0e0; border-radius: 8px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.04); overflow: hidden; }
        .setup-order-box { background: #f3eee6; width: 65px; display: flex; flex-direction: column; align-items: center; justify-content: center; border-right: 1px solid #e0e0e0; padding: 15px 5px; }
        .setup-order-select { font-size: 24px; font-weight: 900; font-family: "Arial Black", sans-serif; border: none; background: transparent; text-align: center; cursor: pointer; outline: none; appearance: none; -webkit-appearance: none; color: #333; }
        .setup-order-label { font-size: 11px; font-weight: bold; color: #555; margin-top: -2px; }
        .setup-order-icon { font-size: 18px; color: #999; margin-top: 5px; font-weight: bold; }
        .setup-details { flex: 1; padding: 15px; }
        .setup-name { font-size: 19px; font-weight: 900; color: #222; margin-bottom: 2px; }
        .setup-meta { font-size: 12px; color: #777; margin-bottom: 12px; font-weight: bold; }
        .setup-control-row { display: flex; align-items: center; gap: 10px; }
        .setup-pos-select { flex: 1; padding: 10px; font-size: 15px; font-weight: bold; border: 1px solid #dcdcdc; border-radius: 6px; background: #fff; color: #333; outline: none; }
        .setup-badge { padding: 6px 10px; font-size: 13px; font-weight: 900; border-radius: 4px; white-space: nowrap; }

        /* タブ・テーブル（結果画面用） */
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
        th { background: #eee; }
        .tab-group { display: flex; margin-bottom: 15px; border-bottom: 2px solid #ddd; }
        .tab { padding: 10px 20px; cursor: pointer; font-weight: bold; color: #777; transition: 0.2s; }
        .tab.active { color: #1976D2; border-bottom: 3px solid #1976D2; }
        .tab-content { display: none; overflow-x: auto; }
        .tab-content.active { display: block; }
    </style>
</head>
<body>
    <div class="container">
        <!-- 1. ローディング -->
        <div id="screen-loading" class="screen active">
            <div id="loading-text">選手データを読み込み中...</div>
        </div>

        <!-- 2. ドラフト画面 -->
        <div id="screen-draft" class="screen">
            <div class="header">
                <div class="draft-info">
                    <span id="draft-type-label">野手ドラフト</span> 
                    <span id="draft-progress" style="margin-left:10px; color:#1976D2;">0 / 9</span>
                </div>
                <div id="draft-skips" style="font-size:14px; font-weight:bold; color:#757575;">見送り残り：5回</div>
            </div>
            <div id="draft-card"></div>
            <div class="btn-group">
                <button class="btn-take" onclick="takePlayer()">獲得する</button>
                <button id="btn-skip" class="btn-skip" onclick="skipPlayer()">見送る</button>
            </div>
        </div>

        <!-- 3. 起用設定画面 (写真風デザイン) -->
        <div id="screen-setup" class="screen">
            <h2 style="text-align:center; margin-bottom:5px;">起用設定</h2>
            <p style="text-align:center; color:#666; font-size:13px; margin-top:0;">打順を変更すると自動で入れ替わります</p>
            
            <h3 style="border-bottom: 2px solid #00796B; color: #00796B; padding-bottom:5px;">スタメン・打順</h3>
            <div id="setup-fielders"></div>
            
            <h3 style="border-bottom: 2px solid #E53935; color: #E53935; margin-top:30px; padding-bottom:5px;">投手ローテーション</h3>
            <div id="setup-pitchers"></div>
            
            <button class="btn-primary" onclick="startPennant()">この設定でペナントレースを開始！</button>
        </div>

        <!-- 4. 結果画面 -->
        <div id="screen-result" class="screen">
            <h2 style="text-align:center;">ペナントレース結果</h2>
            <div id="final-rank" style="text-align:center; font-size:24px; font-weight:bold; margin-bottom:20px;"></div>
            
            <div class="tab-group">
                <div class="tab active" onclick="switchTab('tab-standings')">順位表</div>
                <div class="tab" onclick="switchTab('tab-fielders')">野手成績</div>
                <div class="tab" onclick="switchTab('tab-pitchers')">投手成績</div>
            </div>

            <div id="tab-standings" class="tab-content active"></div>
            <div id="tab-fielders" class="tab-content"></div>
            <div id="tab-pitchers" class="tab-content"></div>

            <button class="btn-primary" onclick="location.reload()" style="background-color:#757575; margin-top:30px;">最初からやり直す</button>
        </div>
    </div>

<script>
let fieldersData = [];
let pitchersData = [];
let draftPool = [];
let myTeam = { fielders: [], pitchers: [] };

let currentDraftType = 'fielder';
let currentCandidate = null;
let requiredCount = 9;
let skipsLeft = 5;

const POSITIONS = ['指名打者', '捕手', '一塁手', '二塁手', '三塁手', '遊撃手', '左翼手', '中堅手', '右翼手'];

// ==========================================
// 1. データ読み込み ＆ 共通関数
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
        console.error(e);
        const loadTxt = document.getElementById('loading-text');
        if (loadTxt) loadTxt.innerText = "読み込みエラー。ファイルを確認してください。";
    }
});

function getRankInfo(num) {
    if (isNaN(num)) return { rank: '-', color: '#999' };
    if (num >= 90) return { rank: 'S', color: '#D4AF37' };
    if (num >= 80) return { rank: 'A', color: '#4CAF50' }; // 緑色系に変更（写真参考）
    if (num >= 70) return { rank: 'B', color: '#FF9800' };
    if (num >= 60) return { rank: 'C', color: '#FBC02D' };
    if (num >= 50) return { rank: 'D', color: '#E53935' };
    if (num >= 40) return { rank: 'E', color: '#7CB342' };
    if (num >= 20) return { rank: 'F', color: '#03A9F4' };
    return { rank: 'G', color: '#9E9E9E' };
}

// 守備位置の日本語変換と適性取得
function parsePositionName(shortPos) {
    const map = {'捕':'捕手','一':'一塁手','二':'二塁手','三':'三塁手','遊':'遊撃手','左':'左翼手','中':'中堅手','右':'右翼手','DH':'指名打者'};
    return map[shortPos] || shortPos;
}

function getPositionAptitude(player, targetPos) {
    if (targetPos === '指名打者') return 90; // DHは高適性扱い
    if (!player["守備力"]) return 20; 
    if (typeof player["守備力"] === 'number') return player["守備力"];
    
    let apt = 10;
    let parts = player["守備力"].toString().split('/');
    parts.forEach(p => {
        let [pos, val] = p.split(':');
        if (pos && parsePositionName(pos.trim()) === targetPos) {
            apt = parseInt(val, 10);
        }
    });
    return apt;
}

// ==========================================
// 2. ドラフト進行処理（省略・前回のまま）
// ==========================================
function showScreen(id) {
    document.querySelectorAll('.screen').forEach(el => el.classList.remove('active'));
    document.getElementById(id).classList.add('active');
}
function updateDraftHeader() { /* ...（前回のコード）... */
    const currentCount = currentDraftType === 'fielder' ? myTeam.fielders.length : myTeam.pitchers.length;
    const remain = requiredCount - currentCount;
    document.getElementById('draft-type-label').innerText = currentDraftType === 'fielder' ? "野手ドラフト" : "投手ドラフト";
    document.getElementById('draft-progress').innerHTML = `${currentCount} / ${requiredCount} <span style="font-size:14px; color:#666;">あと${remain}人</span>`;
    const skipBtn = document.getElementById('btn-skip');
    if (skipsLeft > 0) {
        document.getElementById('draft-skips').innerText = `見送り残り：${skipsLeft}回`;
        skipBtn.classList.remove('btn-disabled'); skipBtn.disabled = false;
    } else {
        document.getElementById('draft-skips').innerText = `見送り不可`;
        skipBtn.classList.add('btn-disabled'); skipBtn.disabled = true;
    }
}
function startDraft(type) {
    currentDraftType = type; requiredCount = type === 'fielder' ? 9 : 15; skipsLeft = 5;
    draftPool = type === 'fielder' ? [...fieldersData] : [...pitchersData];
    showScreen('screen-draft'); nextCandidate();
}
function nextCandidate() {
    const currentCount = currentDraftType === 'fielder' ? myTeam.fielders.length : myTeam.pitchers.length;
    if (currentCount >= requiredCount) {
        if (currentDraftType === 'fielder') startDraft('pitcher');
        else initSetupScreen();
        return;
    }
    if (draftPool.length === 0) { alert("候補枯渇！"); return; }
    const randIdx = Math.floor(Math.random() * draftPool.length);
    currentCandidate = draftPool[randIdx];
    renderCard(currentCandidate, currentDraftType);
    updateDraftHeader();
}
function takePlayer() {
    if (currentDraftType === 'fielder') myTeam.fielders.push(currentCandidate);
    else myTeam.pitchers.push(currentCandidate);
    draftPool = draftPool.filter(p => p !== currentCandidate); nextCandidate();
}
function skipPlayer() {
    if (skipsLeft > 0) { skipsLeft--; draftPool = draftPool.filter(p => p !== currentCandidate); nextCandidate(); }
}
function renderCard(player, type) { /* ...（前回のコード・省略せずにそのまま機能）... */
    const name = player["選手名"] || "不明";
    const team = player["チーム"] || "フリー";
    const htmlContainer = document.getElementById('draft-card');
    
    function createStatCard(label, val) {
        const num = parseInt(val, 10);
        const data = getRankInfo(num);
        return `<div class="stat-card"><div class="stat-label">${label}</div><div class="stat-rank" style="color:${data.color};">${data.rank}</div><div class="stat-val">${isNaN(num)?'-':num}</div></div>`;
    }

    if (type === 'fielder') {
        htmlContainer.innerHTML = `<div class="card"><div class="p-tag">野手</div><div class="p-name">${name}</div><div class="p-meta">${team}</div><div class="stats-box grid-4">${createStatCard('ミート', player["ミート"])}${createStatCard('パワー', player["パワー"])}${createStatCard('走力', player["走力"])}${createStatCard('守備力', player["守備力"]?player["守備力"].split(':')[1]||50:50)}</div></div>`;
    } else {
        htmlContainer.innerHTML = `<div class="card"><div class="p-tag" style="background-color:#E53935;">投手</div><div class="p-name">${name}</div><div class="p-meta">${team}</div><div class="stats-box grid-3">${createStatCard('球威', player["球威"])}${createStatCard('制球', player["制球"])}${createStatCard('スタミナ', player["スタミナ"])}</div></div>`;
    }
}

// ==========================================
// 3. 写真風・起用設定画面（打順自動入れ替え対応）
// ==========================================
function initSetupScreen() {
    showScreen('screen-setup');
    myTeam.fielders.forEach((p, i) => { if (!p.order) p.order = i + 1; });
    renderSetupTables();
}

function renderSetupTables() {
    // 打順でソート
    myTeam.fielders.sort((a, b) => a.order - b.order);

    // 野手UI生成
    let fHtml = '';
    myTeam.fielders.forEach((p, index) => {
        let orderOptions = '';
        for(let i=1; i<=myTeam.fielders.length; i++) {
            orderOptions += `<option value="${i}" ${p.order === i ? 'selected' : ''}>${i}</option>`;
        }

        // 初期ポジションの取得
        let rawDef = p["守備力"] ? p["守備力"].split(':')[0] : 'DH';
        let defaultPos = p.pos || parsePositionName(rawDef);
        if(!p.pos) p.pos = defaultPos; 

        // 適性バッジの計算
        let aptitude = getPositionAptitude(p, p.pos);
        let rankInfo = getRankInfo(aptitude);
        let badgeHtml = `<span class="setup-badge" style="background-color:${rankInfo.color}15; color:${rankInfo.color}; border: 1px solid ${rankInfo.color};">適性${rankInfo.rank}</span>`;

        let posOptions = '';
        POSITIONS.forEach(pos => {
            let r = getRankInfo(getPositionAptitude(p, pos)).rank;
            posOptions += `<option value="${pos}" ${p.pos === pos ? 'selected' : ''}>${pos}（適性${r}）</option>`;
        });

        // 年齢出身の代わりに能力表示
        let meet = parseInt(p["ミート"])||50; let pow = parseInt(p["パワー"])||50; let spd = parseInt(p["走力"])||50;

        fHtml += `
        <div class="setup-card">
            <div class="setup-order-box">
                <select class="setup-order-select" onchange="swapOrder('fielder', ${index}, this.value)">${orderOptions}</select>
                <div class="setup-order-label">番</div>
                <div class="setup-order-icon">≡</div>
            </div>
            <div class="setup-details">
                <div class="setup-name">${p["選手名"]}</div>
                <div class="setup-meta">ミート:${meet} パワー:${pow} 走力:${spd} 本職:${parsePositionName(rawDef)}</div>
                <div class="setup-control-row">
                    <select class="setup-pos-select" onchange="updatePos(${index}, this.value)">${posOptions}</select>
                    ${badgeHtml}
                </div>
            </div>
        </div>`;
    });
    document.getElementById('setup-fielders').innerHTML = fHtml;

    // 投手UI生成
    let pHtml = '';
    const defaultPitcherRoles = ['先発','先発','先発','先発','先発','先発','中継ぎエース','中継ぎ','中継ぎ','中継ぎ','中継ぎ','敗戦処理','抑え'];
    myTeam.pitchers.forEach((p, index) => {
        if (!p.role) p.role = defaultPitcherRoles[index] || '中継ぎ';
        
        let roleOptions = '';
        ['先発', '中継ぎエース', '中継ぎ', 'ビハインド', '敗戦処理', '抑え'].forEach(r => {
            roleOptions += `<option value="${r}" ${p.role === r ? 'selected' : ''}>${r}</option>`;
        });

        let kyui = parseInt(p["球威"])||50; let ctrl = parseInt(p["制球"])||50; let sta = parseInt(p["スタミナ"])||50;
        let rankInfo = getRankInfo(sta); // スタミナを適性の代わりとして表示

        pHtml += `
        <div class="setup-card">
            <div class="setup-order-box" style="background:#f9f9f9;">
                <div class="setup-order-select" style="cursor:default;">${index + 1}</div>
                <div class="setup-order-label">番手</div>
                <div class="setup-order-icon">≡</div>
            </div>
            <div class="setup-details">
                <div class="setup-name">${p["選手名"]}</div>
                <div class="setup-meta">球威:${kyui} 制球:${ctrl} スタミナ:${sta}</div>
                <div class="setup-control-row">
                    <select class="setup-pos-select" onchange="updatePitcherRole(${index}, this.value)">${roleOptions}</select>
                    <span class="setup-badge" style="background-color:${rankInfo.color}15; color:${rankInfo.color}; border: 1px solid ${rankInfo.color};">ｽﾀﾐﾅ${rankInfo.rank}</span>
                </div>
            </div>
        </div>`;
    });
    document.getElementById('setup-pitchers').innerHTML = pHtml;
}

// 打順の自動入れ替え処理
function swapOrder(type, currentIndex, newValue) {
    const targetOrder = parseInt(newValue, 10);
    const teamArray = type === 'fielder' ? myTeam.fielders : myTeam.pitchers;
    const currentOrder = teamArray[currentIndex].order;

    const swapTarget = teamArray.find(p => p.order === targetOrder && p !== teamArray[currentIndex]);
    if (swapTarget) swapTarget.order = currentOrder;
    teamArray[currentIndex].order = targetOrder;
    
    renderSetupTables(); // UIを再描画して順番を並び替え
}

function updatePos(index, newPos) { myTeam.fielders[index].pos = newPos; renderSetupTables(); }
function updatePitcherRole(index, newRole) { myTeam.pitchers[index].role = newRole; }


// ==========================================
// 4. ペナント実行 ＆ Python完全移植シミュレーション
// ==========================================
const TEAM_CLUTCH_MODIFIERS = { "阪神": 0.0, "広島": 0.0, "DeNA": 10.0, "巨人": 0.0, "ヤクルト": 40.0, "中日": -90.0, "オリックス": 0.0, "ロッテ": 0.0, "ソフトバンク": -30.0, "楽天": 0.0, "西武": 0.0, "日本ハム": 0.0, "フリー": 0.0, "マイチーム": 0.0 };

function clamp(x, lo, hi) { return Math.max(lo, Math.min(hi, x)); }
function fatigueFactor(stamina, gameOuts) {
    let innings = gameOuts / 3.0; if (innings <= 4) return 1.0;
    let excess = innings - 4; let penalty = excess * Math.max(0.0, (80.0 - stamina)) * 0.0015;
    return clamp(1.0 - penalty, 0.82, 1.0);
}

function atBatProbabilities(batter, pitcher, gameOuts) {
    let bCon = parseInt(batter["ミート"], 10) || 50; let bPow = parseInt(batter["パワー"], 10) || 50;
    let pPow = parseInt(pitcher["球威"], 10) || 50; let pCtrl = parseInt(pitcher["制球"], 10) || 50;
    let pStamina = parseInt(pitcher["スタミナ"], 10) || 50;
    
    let pitchQuality = 50.0; let pitchVariety = 1;
    let soQuality = (pitchQuality * 0.4) + (pPow * 0.4) + (pCtrl * 0.2) + 5.0;
    let bbQuality = (pCtrl * 0.8) + (pitchQuality * 0.2) + 5.0;

    let fatigue = fatigueFactor(pStamina, gameOuts);
    let bbDiff = bbQuality - 50.0; let rawContactDiff = bCon - soQuality;

    let contactPenalty = 0.0; if (bCon < 45.0) { let diff = 45.0 - bCon; contactPenalty = (diff * 0.4) + (Math.pow(diff, 2) * 0.01); }
    let powerPenalty = 0.0; if (bPow < 60.0) { let diff = 60.0 - bPow; powerPenalty = (diff * 0.4) + (Math.pow(diff, 2) * 0.01); }

    let walkBonus = 0.0;
    if (bPow > 50.0) walkBonus += (bPow - 50.0) * 0.0012;
    if (bCon > 50.0) walkBonus += (bCon - 50.0) * 0.0006;
    if (bPow >= 80.0) walkBonus += (bPow - 80.0) * 0.0015;

    let walk = 0.095 - (bbDiff * 0.0015) + walkBonus;
    let so = 0.155 - (rawContactDiff * 0.0008) + (contactPenalty * 0.0015) + (powerPenalty * 0.0008) + (soQuality - 60.0) * 0.0055;

    if (fatigue < 1.0) { walk += (1.0 - fatigue) * 0.02; so -= (1.0 - fatigue) * 0.03; }
    walk = clamp(walk, 0.010, 0.20); so = clamp(so, 0.05, 0.45);
    let inPlay = Math.max(0.20, 1.0 - (walk + so));

    let tot = walk + so + inPlay;
    return [{ result: "walk", prob: walk / tot }, { result: "so", prob: so / tot }, { result: "in_play", prob: inPlay / tot }];
}

function resolveStatcastInPlay(batter, pitcher, ability) {
    let bCon = parseInt(batter["ミート"], 10) || 50; let bPow = parseInt(batter["パワー"], 10) || 50;
    let pPow = parseInt(pitcher["球威"], 10) || 50; let pCtrl = parseInt(pitcher["制球"], 10) || 50;

    let gbProb = 0.45, fbProb = 0.25, ldProb = 0.20, puProb = 0.10;
    if (bPow > 50.0) { fbProb += (bPow - 50.0) * 0.0015; gbProb -= (bPow - 50.0) * 0.0015; }
    if (bCon > 50.0) { ldProb += (bCon - 50.0) * 0.0025; puProb -= (bCon - 50.0) * 0.0025; }

    gbProb += (pCtrl - 50.0) * 0.0020; ldProb -= (pCtrl - 50.0) * 0.0020;
    puProb += (pPow - 50.0) * 0.0020; fbProb -= (pPow - 50.0) * 0.0020;

    gbProb = Math.max(0.05, gbProb); fbProb = Math.max(0.05, fbProb); ldProb = Math.max(0.05, ldProb); puProb = Math.max(0.01, puProb);

    let tot = gbProb + fbProb + ldProb + puProb; let r = Math.random() * tot;
    let bType = r < gbProb ? "GB" : (r < gbProb + fbProb ? "FB" : (r < gbProb + fbProb + ldProb ? "LD" : "PU"));

    let baseErr = 0.02; let outcome = "out";
    if (bType === "GB") {
        if (Math.random() > clamp(0.72 + (ability - 55.0) * 0.006, 0.40, 0.95)) outcome = Math.random() < 0.10 ? "double" : "single";
        else outcome = Math.random() < clamp(baseErr * (1.0 + (55.0 - ability)/40.0), baseErr*0.2, baseErr*3.0) ? "error" : "out";
    } else if (bType === "FB") {
        let hrProb = 0.0; let p = bPow;
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

        if (Math.random() < clamp(hrProb, 0.0, 0.40)) outcome = "hr";
        else {
            if (Math.random() > clamp(0.82 + (ability - 55.0)*0.005, 0.40, 0.99)) {
                let hR = Math.random(); outcome = hR < 0.65 ? "single" : (hR < 0.95 ? "double" : "triple");
            } else outcome = Math.random() < clamp(baseErr * (1.0 + (55.0 - ability)/40.0), baseErr*0.2, baseErr*3.0) ? "error" : "out";
        }
    } else if (bType === "LD") {
        if (Math.random() > clamp(0.30 + (ability - 55.0)*0.005, 0.10, 0.60)) {
            let hR = Math.random(); outcome = hR < 0.70 ? "single" : (hR < 0.95 ? "double" : "triple");
        } else outcome = Math.random() < clamp(baseErr * (1.0 + (55.0 - ability)/40.0), baseErr*0.2, baseErr*3.0) ? "error" : "out";
    } else {
        if (Math.random() > clamp(0.98 + (ability - 55.0)*0.002, 0.60, 1.0)) outcome = "single";
        else outcome = Math.random() < clamp(baseErr * (1.0 + (55.0 - ability)/40.0), baseErr*0.2, baseErr*3.0) ? "error" : "out";
    }
    return outcome;
}

function simulateAtBatOutcome(batter, pitcher, defMax, gameOuts) {
    let probs = atBatProbabilities(batter, pitcher, gameOuts);
    let r1 = Math.random(); let cum = 0.0; let step1 = "in_play";
    for (let p of probs) { cum += p.prob; if (r1 <= cum) { step1 = p.result; break; } }
    if (step1 === "so" || step1 === "walk") return step1;
    return resolveStatcastInPlay(batter, pitcher, defMax);
}

function advanceOnHit(bases, batter, result) {
    let newBases = [null, null, null]; let runs = 0;
    if (result === "hr") { runs += (bases[0]?1:0)+(bases[1]?1:0)+(bases[2]?1:0)+1; return { newBases, runs }; }
    if (result === "triple") { runs += (bases[0]?1:0)+(bases[1]?1:0)+(bases[2]?1:0); newBases[2] = batter; return { newBases, runs }; }
    if (result === "double") {
        if (bases[2]) runs++;
        if (bases[1]) { if (Math.random() < 0.75) runs++; else newBases[2] = bases[1]; }
        if (bases[0]) { if (Math.random() < 0.50) runs++; else newBases[2] = bases[0]; }
        newBases[1] = batter; return { newBases, runs };
    }
    if (bases[2]) runs++;
    if (bases[1]) { if (Math.random() < 0.75) runs++; else newBases[2] = bases[1]; }
    if (bases[0]) { if (Math.random() < 0.40) newBases[2] = bases[0]; else newBases[1] = bases[0]; }
    newBases[0] = batter; return { newBases, runs };
}
function advanceOnWalk(bases, batter) {
    let newBases = [...bases]; let runs = 0;
    if (bases[0] && bases[1] && bases[2]) runs = 1;
    if (bases[0] && bases[1]) newBases[2] = bases[1];
    if (bases[0]) newBases[1] = bases[0];
    newBases[0] = batter; return { newBases, runs };
}

function switchTab(tabId) {
    document.querySelectorAll('.tab').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    event.target.classList.add('active'); document.getElementById(tabId).classList.add('active');
}

function startPennant() {
    myTeam.fielders.forEach(p => p.stats = { pa: 0, ab: 0, hit: 0, double: 0, triple: 0, hr: 0, bb: 0, so: 0, rbi: 0, runs: 0 });
    myTeam.pitchers.forEach(p => p.stats = { games: 0, outs: 0, runs: 0, er: 0, wins: 0, losses: 0, saves: 0, holds: 0, so: 0, bb: 0, hr: 0, hits: 0 });

    showScreen('screen-result');
    
    let teams = [
        { name: "あなたのオリジナルチーム", wins: 0, losses: 0, strength: 50 },
        { name: "ライバル球団A", wins: 0, losses: 0, strength: 55, def: 55 },
        { name: "ライバル球団B", wins: 0, losses: 0, strength: 50, def: 50 },
        { name: "ライバル球団C", wins: 0, losses: 0, strength: 50, def: 50 },
        { name: "ライバル球団D", wins: 0, losses: 0, strength: 45, def: 45 },
        { name: "ライバル球団E", wins: 0, losses: 0, strength: 40, def: 40 }
    ];

    let starters = myTeam.pitchers.filter(p => p.role === '先発');
    let relievers = myTeam.pitchers.filter(p => p.role !== '先発' && p.role !== '抑え');
    let closers = myTeam.pitchers.filter(p => p.role === '抑え');
    if (starters.length === 0) starters = myTeam.pitchers;

    for (let g = 0; g < 143; g++) {
        let enemy = teams[Math.floor(Math.random() * 5) + 1];
        let myScore = 0; let enemyScore = 0;
        let currentBatterIdx = 0;
        let enemyPitcher = { "球威": enemy.strength, "制球": enemy.strength, "スタミナ": 60 };
        let currentPitcher = starters[g % starters.length];
        currentPitcher.stats.games++;
        let pitcherStamina = parseInt(currentPitcher["スタミナ"], 10) || 50;
        let pitchesThrown = 0;

        for (let inning = 1; inning <= 9; inning++) {
            let outs = 0; let bases = [null, null, null];
            while (outs < 3) {
                let baseBatter = myTeam.fielders[currentBatterIdx % myTeam.fielders.length];
                let calcBatter = Object.assign({}, baseBatter);
                let isScoringPos = (bases[1] !== null || bases[2] !== null);
                if (isScoringPos) {
                    let clutchMod = TEAM_CLUTCH_MODIFIERS[calcBatter["チーム"] || "フリー"] || 0.0;
                    calcBatter["ミート"] = Math.max(1.0, (parseInt(calcBatter["ミート"], 10) || 50) + clutchMod);
                    calcBatter["パワー"] = Math.max(1.0, (parseInt(calcBatter["パワー"], 10) || 50) + clutchMod);
                }
                baseBatter.stats.pa++;
                let pitcherOuts = (inning - 1) * 3 + outs;
                let result = simulateAtBatOutcome(calcBatter, enemyPitcher, enemy.def, pitcherOuts);

                if (result === "so") { baseBatter.stats.ab++; baseBatter.stats.so++; outs++; } 
                else if (result === "out" || result === "error") { baseBatter.stats.ab++; outs++; } 
                else if (result === "walk") {
                    baseBatter.stats.bb++;
                    let adv = advanceOnWalk(bases, baseBatter); bases = adv.newBases;
                    if (adv.runs > 0) { myScore += adv.runs; baseBatter.stats.rbi += adv.runs; }
                } else {
                    baseBatter.stats.ab++; baseBatter.stats.hit++;
                    if (result === "double") baseBatter.stats.double++;
                    if (result === "triple") baseBatter.stats.triple++;
                    if (result === "hr") baseBatter.stats.hr++;
                    let adv = advanceOnHit(bases, baseBatter, result); bases = adv.newBases;
                    if (adv.runs > 0) { myScore += adv.runs; baseBatter.stats.rbi += adv.runs; }
                }
                currentBatterIdx++;
            }

            let enemyOuts = 0; let enemyBases = [null, null, null];
            if (inning >= 9 && myScore > enemyScore && myScore - enemyScore <= 3 && closers.length > 0 && currentPitcher !== closers[0]) {
                currentPitcher = closers[0]; currentPitcher.stats.games++; pitchesThrown = 0;
            } else if ((pitchesThrown > pitcherStamina * 1.5 || currentPitcher.stats.runs >= 4) && relievers.length > 0 && inning < 9) {
                currentPitcher = relievers[Math.floor(Math.random() * relievers.length)]; currentPitcher.stats.games++; pitchesThrown = 0;
            }

            while (enemyOuts < 3) {
                pitchesThrown += Math.floor(Math.random() * 5) + 3;
                let enemyBatter = { "ミート": enemy.strength, "パワー": enemy.strength, "走力": enemy.strength };
                let eResult = simulateAtBatOutcome(enemyBatter, currentPitcher, 50, currentPitcher.stats.outs);
                
                if (eResult === "so") { currentPitcher.stats.so++; enemyOuts++; currentPitcher.stats.outs++; } 
                else if (eResult === "out" || eResult === "error") { enemyOuts++; currentPitcher.stats.outs++; } 
                else if (eResult === "walk") {
                    currentPitcher.stats.bb++;
                    let adv = advanceOnWalk(enemyBases, enemyBatter); enemyBases = adv.newBases;
                    if (adv.runs > 0) { enemyScore += adv.runs; currentPitcher.stats.runs += adv.runs; currentPitcher.stats.er += adv.runs; }
                } else {
                    currentPitcher.stats.hits++; if (eResult === "hr") currentPitcher.stats.hr++;
                    let adv = advanceOnHit(enemyBases, enemyBatter, eResult); enemyBases = adv.newBases;
                    if (adv.runs > 0) { enemyScore += adv.runs; currentPitcher.stats.runs += adv.runs; currentPitcher.stats.er += adv.runs; }
                }
                if (inning >= 9 && enemyScore > myScore) break;
            }
            if (inning >= 9 && enemyScore > myScore) break;
        }

        if (myScore > enemyScore) {
            teams[0].wins++; enemy.losses++;
            starters[g % starters.length].stats.wins++;
            if (currentPitcher.role === '抑え' && myScore - enemyScore <= 3) currentPitcher.stats.saves++;
        } else if (myScore < enemyScore) {
            teams[0].losses++; enemy.wins++;
            currentPitcher.stats.losses++;
        }
    }

    // --- 順位表 ---
    teams.sort((a, b) => b.wins - a.wins);
    document.getElementById('final-rank').innerHTML = `6チーム中 <span style="color:#c62828; font-size:40px;">${teams.findIndex(t => t.name === "あなたのオリジナルチーム") + 1}位</span>`;
    let tableHtml = `<table><tr><th>順位</th><th>チーム名</th><th>勝利</th><th>敗北</th><th>勝率</th></tr>`;
    teams.forEach((t, index) => {
        let wp = (t.wins / (t.wins + t.losses)).toFixed(3).replace(/^0\./, '.');
        let isMyTeam = t.name === "あなたのオリジナルチーム" ? 'style="background-color:#fff9c4; font-weight:bold;"' : '';
        tableHtml += `<tr ${isMyTeam}><td>${index+1}</td><td>${t.name}</td><td>${t.wins}</td><td>${t.losses}</td><td>${wp}</td></tr>`;
    });
    document.getElementById('tab-standings').innerHTML = tableHtml + `</table>`;

    // --- 野手成績 ---
    let statFHtml = `<table><tr><th>打順</th><th>守備位置</th><th>選手名</th><th>打率</th><th>本塁打</th><th>打点</th><th>OPS</th></tr>`;
    myTeam.fielders.forEach(p => {
        let avg = p.stats.ab > 0 ? p.stats.hit / p.stats.ab : 0;
        let obp = p.stats.pa > 0 ? (p.stats.hit + p.stats.bb) / p.stats.pa : 0;
        let tb = (p.stats.hit - p.stats.double - p.stats.triple - p.stats.hr) + (p.stats.double * 2) + (p.stats.triple * 3) + (p.stats.hr * 4);
        let slg = p.stats.ab > 0 ? tb / p.stats.ab : 0;
        statFHtml += `<tr><td>${p.order}</td><td>${p.pos}</td><td style="font-weight:bold;">${p["選手名"]}</td><td>${avg.toFixed(3).replace(/^0\./, '.')}</td><td>${p.stats.hr}</td><td>${p.stats.rbi}</td><td>${(obp + slg).toFixed(3).replace(/^0\./, '.')}</td></tr>`;
    });
    document.getElementById('tab-fielders').innerHTML = statFHtml + `</table>`;

    // --- 投手成績 ---
    let statPHtml = `<table><tr><th>役割</th><th>選手名</th><th>防御率</th><th>登板</th><th>投球回</th><th>勝</th><th>敗</th><th>セ</th><th>奪三振</th></tr>`;
    myTeam.pitchers.forEach(p => {
        let era = p.stats.outs > 0 ? (p.stats.er * 27) / p.stats.outs : 0.0;
        statPHtml += `<tr><td>${p.role}</td><td style="font-weight:bold;">${p["選手名"]}</td><td>${era.toFixed(2)}</td><td>${p.stats.games}</td><td>${Math.floor(p.stats.outs / 3)}.${p.stats.outs % 3}</td><td>${p.stats.wins}</td><td>${p.stats.losses}</td><td>${p.stats.saves}</td><td>${p.stats.so}</td></tr>`;
    });
    document.getElementById('tab-pitchers').innerHTML = statPHtml + `</table>`;
}
</script>
</body>
</html>
