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
window.onload = async () => {
    try {
        // ★ リポジトリルートのファイルを絶対パスで指定
        const res = await fetch('/baseball_data.xlsx');
        if (!res.ok) throw new Error(`HTTPエラー ステータス: ${res.status}`);
        
        const buffer = await res.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        
        // シート名はプロテクト側と同じ「野手」「投手」を指定
        const rawFielders = XLSX.utils.sheet_to_json(workbook.Sheets['野手']) || [];
        const rawPitchers = XLSX.utils.sheet_to_json(workbook.Sheets['投手']) || [];
        
        // 選手名があるデータのみ抽出
        fieldersData = rawFielders.filter(p => p['選手名']);
        pitchersData = rawPitchers.filter(p => p['選手名']);
        
        if (fieldersData.length === 0 || pitchersData.length === 0) {
            throw new Error("シート『野手』または『投手』にデータが見つかりません。");
        }

        startDraft('fielder');
    } catch (e) {
        document.getElementById('loading-text').innerHTML = `<span style="color:#c62828; font-weight:bold;">エラー: ${e.message}</span>`;
        console.error(e);
    }
};

// ==========================================
// 2. 画面切り替え・UI更新
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

// ==========================================
// 3. ドラフト進行ロジック
// ==========================================
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
        if (currentDraftType === 'fielder') {
            startDraft('pitcher');
        } else {
            initSetupScreen();
        }
        return;
    }

    if (draftPool.length === 0) {
        alert("候補選手が枯渇しました！最初からやり直してください。");
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

function renderCard(player, type) {
    const name = player["選手名"] || "不明";
    const team = player["チーム"] || player["コード"] || "フリー";
    const htmlContainer = document.getElementById('draft-card');
    
    // エクセルの列名に合わせて表示内容を調整（存在しない場合は打率などを表示）
    if (type === 'fielder') {
        const contact = player["ミート"] || player["打率"] || "-";
        const power = player["パワー"] || player["本塁打"] || "-";
        const rbi = player["打点"] || "-";

        htmlContainer.innerHTML = `
            <div class="card">
                <div class="p-tag">野手</div>
                <div class="p-name">${name}</div>
                <div class="p-meta">${team}</div>
                <div class="stats-box">
                    <div class="stat-item"><div class="stat-label">ミート/打率</div><div class="stat-val">${contact}</div></div>
                    <div class="stat-item"><div class="stat-label">パワー/HR</div><div class="stat-val">${power}</div></div>
                    <div class="stat-item"><div class="stat-label">打点</div><div class="stat-val">${rbi}</div></div>
                </div>
            </div>
        `;
    } else {
        const era = player["防御率"] || player["球威"] || "-";
        const win = player["勝利"] || "-";
        const strike = player["奪三振"] || "-";

        htmlContainer.innerHTML = `
            <div class="card">
                <div class="p-tag" style="background-color: #E53935;">投手</div>
                <div class="p-name">${name}</div>
                <div class="p-meta">${team}</div>
                <div class="stats-box">
                    <div class="stat-item"><div class="stat-label">防御率/球威</div><div class="stat-val">${era}</div></div>
                    <div class="stat-item"><div class="stat-label">勝利</div><div class="stat-val">${win}</div></div>
                    <div class="stat-item"><div class="stat-label">奪三振</div><div class="stat-val">${strike}</div></div>
                </div>
            </div>
        `;
    }
}

// ==========================================
// 4. オーダー確認 ＆ ペナント実行
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
    
    // 架空のペナントシミュレーション（UI確認用・順位計算）
    // 本格的な成績計算ロジックはPythonからさらに移植していく土台になります
    let teams = [
        { name: "あなたのオリジナルチーム", wins: 0, losses: 0 },
        { name: "ライバル球団A", wins: 0, losses: 0 },
        { name: "ライバル球団B", wins: 0, losses: 0 },
        { name: "ライバル球団C", wins: 0, losses: 0 },
        { name: "ライバル球団D", wins: 0, losses: 0 },
        { name: "ライバル球団E", wins: 0, losses: 0 }
    ];

    // あなたのチームの強さを能力値（今回は簡易乱数）で算出
    let baseStrength = 0.45 + (Math.random() * 0.15); 

    // 143試合を消化
    for(let i=0; i<143; i++) {
        if(Math.random() < baseStrength) teams[0].wins++;
        else teams[0].losses++;
        
        for(let j=1; j<6; j++) {
            if(Math.random() < 0.5) teams[j].wins++;
            else teams[j].losses++;
        }
    }

    // 順位ソート
    teams.sort((a, b) => b.wins - a.wins);

    // 結果出力
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
