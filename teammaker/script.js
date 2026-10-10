let fieldersData = [];
let pitchersData = [];
let myTeam = { fielders: [], pitchers: [] };
let draftPool = [];
let currentDraftType = 'fielder'; 
let currentCandidate = null;

// ① 画面読み込み時に、上の階層(../data/)からエクセルを自動取得
window.onload = async () => {
    try {
        const res = await fetch('../data/選手データ.xlsx'); // ★ご自身のファイル名に合わせる
        const buffer = await res.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: 'array' });
        
        // ★シート名をご自身のファイルに合わせて変更してください
        fieldersData = XLSX.utils.sheet_to_json(workbook.Sheets['野手能力']);
        pitchersData = XLSX.utils.sheet_to_json(workbook.Sheets['投手能力']);
        
        console.log("データ読み込み完了！", fieldersData, pitchersData);
        startDraft('fielder');
    } catch (e) {
        document.getElementById('screen-loading').innerHTML = `<p style="color:red; text-align:center;">データの読み込みに失敗しました。パスやファイル名を確認してください。</p>`;
        console.error(e);
    }
};

function showScreen(id) {
    document.querySelectorAll('.screen').forEach(el => el.classList.remove('active'));
    document.getElementById(id).classList.add('active');
}

// ② ドラフト処理
function startDraft(type) {
    currentDraftType = type;
    draftPool = type === 'fielder' ? [...fieldersData] : [...pitchersData];
    showScreen('screen-draft');
    nextCandidate();
}

function nextCandidate() {
    let limit = currentDraftType === 'fielder' ? 9 : 15;
    let currentCount = currentDraftType === 'fielder' ? myTeam.fielders.length : myTeam.pitchers.length;

    if (currentCount >= limit) {
        if (currentDraftType === 'fielder') {
            startDraft('pitcher');
        } else {
            showScreen('screen-setup');
        }
        return;
    }

    let randIdx = Math.floor(Math.random() * draftPool.length);
    currentCandidate = draftPool[randIdx];
    let name = currentCandidate["選手名_出力"] || currentCandidate["選手名"];
    
    document.getElementById('draft-card').innerHTML = `
        <h3>${name}</h3>
        <p>${currentDraftType === 'fielder' ? '野手' : '投手'}</p>
    `;
}

function takePlayer() {
    if (currentDraftType === 'fielder') myTeam.fielders.push(currentCandidate);
    else myTeam.pitchers.push(currentCandidate);
    draftPool = draftPool.filter(p => p !== currentCandidate);
    nextCandidate();
}

function skipPlayer() {
    draftPool = draftPool.filter(p => p !== currentCandidate);
    nextCandidate();
}

// ③ ペナント処理（先ほど作成したシミュレーションエンジンをここで呼び出す）
function startPennant() {
    showScreen('screen-result');
    document.getElementById('result-display').innerHTML = "<p>ペナントレース全日程が終了しました！（※ここに成績表を出力します）</p>";
}
