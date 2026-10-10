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
        const res = await fetch('/fuji210320-takatora.github.io/baseball_data.xlsx');
        if (!res.ok) throw new Error(`HTTPエラー ステータス: ${res.status}`);
        
        const buffer = await res.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        
        // ★修正：プロテクトと同じ「野手」「投手」シートを探す（「野手能力」になっていても対応可能）
        const sheetF = workbook.Sheets['野手'] || workbook.Sheets['野手能力'];
        const sheetP = workbook.Sheets['投手'] || workbook.Sheets['投手能力'];
        
        // それでも見つからない場合のエラー表示（実際にあるシート名を画面に表示させます）
        if (!sheetF || !sheetP) {
            throw new Error(`「野手」または「投手」シートが見つかりません。現在のシート一覧: [ ${workbook.SheetNames.join(', ')} ]`);
        }

        const rawFielders = XLSX.utils.sheet_to_json(sheetF) || [];
        const rawPitchers = XLSX.utils.sheet_to_json(sheetP) || [];
        
        // 「選手名」列があるデータのみ抽出（もし列の名前が「名前」などでも対応）
        fieldersData = rawFielders.filter(p => p['選手名'] || p['名前']);
        pitchersData = rawPitchers.filter(p => p['選手名'] || p['名前']);
        
        if (fieldersData.length === 0 || pitchersData.length === 0) {
            throw new Error("シートは見つかりましたが「選手名」の列が見つかりません。エクセルの1行目を確認してください。");
        }

        // 読み込みが成功したら、ローディング画面を隠してドラフト開始
        document.getElementById('screen-loading').classList.remove('active');
        startDraft('fielder');
    } catch (e) {
        document.getElementById('loading-text').innerHTML = `<span style="color:#c62828; font-weight:bold;">エラー: ${e.message}</span>`;
        console.error(e);
    }
});

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
    const name = player["選手名"] || player["名前"] || "不明";
    const team = player["チーム"] || player["コード"] || "フリー";
    const htmlContainer = document.getElementById('draft-card');
    
    if (type === 'fielder') {
        // ★ご指定通り、ミート・パワー・走力・守備力を表示
        const contact = player["ミート"] || "-";
        const power = player["パワー"] || "-";
        const speed = player["走力"] || "-";
        const defense = player["守備力"] || player["守備"] || "-";

        htmlContainer.innerHTML = `
            <div class="card">
                <div class="p-tag">野手</div>
                <div class="p-name">${name}</div>
                <div class="p-meta">${team}</div>
                <div class="stats-box">
                    <div class="stat-item"><div class="stat-label">ミート</div><div class="stat-val">${contact}</div></div>
                    <div class="stat-item"><div class="stat-label">パワー</div><div class="stat-val">${power}</div></div>
                    <div class="stat-item"><div class="stat-label">走力</div><div class="stat-val">${speed}</div></div>
                    <div class="stat-item"><div class="stat-label">守備力</div><div class="stat-val">${defense}</div></div>
                </div>
            </div>
        `;
    } else {
        const era = player["防御率"] || player["球威"] || player["球速"] || "-";
        const win = player["勝利"] || player["制球"] || player["コントロール"] || "-";
        const strike = player["奪三振"] || player["スタミナ"] || "-";

        htmlContainer.innerHTML = `
            <div class="card">
                <div class="p-tag" style="background-color: #E53935;">投手</div>
                <div class="p-name">${name}</div>
                <div class="p-meta">${team}</div>
                <div class="stats-box">
                    <div class="stat-item"><div class="stat-label">防御率/球威</div><div class="stat-val">${era}</div></div>
                    <div class="stat-item"><div class="stat-label">勝利/制球</div><div class="stat-val">${win}</div></div>
                    <div class="stat-item"><div class="stat-label">奪三振/ｽﾀﾐﾅ</div><div class="stat-val">${strike}</div></div>
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
    
    // 選んだ選手のリストを見やすく表示
    const fList = myTeam.fielders.map(p => (p["選手名"] || p["名前"])).join("、");
    document.getElementById('my-fielders-list').innerText = fList;
    
    const pList = myTeam.pitchers.map(p => (p["選手名"] || p["名前"])).join("、");
    document.getElementById('my-pitchers-list').innerText = pList;
}

function startPennant() {
    showScreen('screen-result');
    
    // 架空のペナントシミュレーション（UI確認用）
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
