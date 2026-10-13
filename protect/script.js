let allPlayersData = [];
let currentTeam = '';
const MAX_PROTECT = 28;
let protectedIds = new Set();
let isDataLoaded = false;

const teamCodeMap = {
    't': '阪神タイガース',
    'yb': '横浜DeNAベイスターズ',
    'g': '読売ジャイアンツ',
    'c': '広島東洋カープ',
    'd': '中日ドラゴンズ',
    's': '東京ヤクルトスワローズ',
    'h': '福岡ソフトバンクホークス',
    'f': '北海道日本ハムファイターズ',
    'bs': 'オリックス・バファローズ',
    'e': '東北楽天ゴールデンイーグルス',
    'l': '埼玉西武ライオンズ',
    'm': '千葉ロッテマリーンズ'
};

const prefectures = ['北海道','青森','岩手','宮城','秋田','山形','福島','茨城','栃木','群馬','埼玉','千葉','東京','神奈川','新潟','富山','石川','福井','山梨','長野','岐阜','静岡','愛知','三重','滋賀','京都','大阪','兵庫','奈良','和歌山','鳥取','島根','岡山','広島','山口','徳島','香川','愛媛','高知','福岡','佐賀','長崎','熊本','大分','宮崎','鹿児島','沖縄'];

function formatAvg(avg) {
    if (avg === undefined || avg === null || avg === '') return '---';
    const num = parseFloat(avg);
    if (isNaN(num)) return '---';
    let str = num.toFixed(3);
    if (str.startsWith('0.')) {
        return str.substring(1); 
    }
    return str;
}

function formatSalary(salary) {
    if (!salary || salary === '-' || salary === 0) return '';
    return String(salary) + ' ';
}

function normalizeName(name) {
    if (!name) return '';
    return String(name).replace(/[\s ]+/g, '');
}

window.addEventListener('DOMContentLoaded', () => {
    
    fetch('/baseball_data.xlsx')
        .then(response => {
            if (!response.ok) throw new Error(`HTTPエラー ステータス: ${response.status}`);
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, {type: 'array'});
            processExcelData(workbook);
            
            isDataLoaded = true;
            document.getElementById('statusMessage').classList.add('hidden');
            document.getElementById('mainHeader').classList.remove('hidden');
            
            const teamSelector = document.getElementById('teamSelector');
            if(teamSelector.options.length > 0) {
                currentTeam = teamSelector.options[0].value;
                updateUI();
            }
        })
        .catch(error => {
            console.error('Error fetching excel:', error);
            // ★修正箇所：フリーズを回避し、画面に具体的なエラー原因を赤字で出力します
            document.getElementById('statusMessage').innerHTML = `<span style="color:red; font-weight:bold;">エラー: ${error.message}</span><br>ファイルパス（リポジトリ名）が間違っているか、ファイルが存在しません。`;
        });
});

function processExcelData(workbook) {
    const basicData = XLSX.utils.sheet_to_json(workbook.Sheets['基本データ']) || [];
    const pitcher1 = XLSX.utils.sheet_to_json(workbook.Sheets['投手']) || [];
    const fielder1 = XLSX.utils.sheet_to_json(workbook.Sheets['野手']) || [];
    const pitcher2 = XLSX.utils.sheet_to_json(workbook.Sheets['2軍投手']) || [];
    const fielder2 = XLSX.utils.sheet_to_json(workbook.Sheets['2軍野手']) || [];

    const createMap = (arr) => {
        let map = {};
        arr.forEach(row => { 
            if (row['選手名']) {
                map[normalizeName(row['選手名'])] = row; 
            }
        });
        return map;
    };
    
    const statsMap = {
        '投手': createMap(pitcher1),
        '野手': createMap(fielder1),
        '2軍投手': createMap(pitcher2),
        '2軍野手': createMap(fielder2)
    };

    let teamsSet = new Set();
    allPlayersData = [];

    basicData.forEach(player => {
        if (!player['選手名'] || !player['コード']) return; 

        const uniformNo = String(player['背番号'] || '').trim();
        if (uniformNo.length >= 3) return;

        const origin = String(player['出身'] || '');
        if (origin) {
            const isJapan = prefectures.some(pref => origin.includes(pref));
            if (!isJapan) return; 
        }

        const code = String(player['コード']).toLowerCase();
        const teamName = teamCodeMap[code] || code;
        teamsSet.add(teamName);

        const position = player['守備位置'];
        const isPitcher = (position === '投手');
        
        const originalName = player['選手名'];
        const normName = normalizeName(originalName);
        
        let stats1Str = '出場なし';
        let stats2Str = '出場なし';

        if (isPitcher) {
            const p1 = statsMap['投手'][normName];
            const p2 = statsMap['2軍投手'][normName];
            
            if (p1) {
                let hp = p1['HP'] ? `${p1['HP']}H ` : '';
                let sv = p1['セーブ'] ? `${p1['セーブ']}S` : '';
                stats1Str = `${p1['登板数']||0}試 ${p1['勝利']||0}勝${p1['敗北']||0}敗 ${hp}${sv}`.trim();
            }
            if (p2) {
                let hp = p2['HP'] ? `${p2['HP']}H ` : '';
                let sv = p2['セーブ'] ? `${p2['セーブ']}S` : '';
                stats2Str = `${p2['登板数']||0}試 ${p2['勝利']||0}勝${p2['敗北']||0}敗 ${hp}${sv}`.trim();
            }
        } else {
            const f1 = statsMap['野手'][normName];
            const f2 = statsMap['2軍野手'][normName];
            
            if (f1) {
                stats1Str = `${f1['試合数']||0}試 ${formatAvg(f1['打率'])} ${f1['本塁打']||0}本 ${f1['打点']||0}点`;
            }
            if (f2) {
                stats2Str = `${f2['試合数']||0}試 ${formatAvg(f2['打率'])} ${f2['本塁打']||0}本 ${f2['打点']||0}点`;
            }
        }

        const statsCombined = `${stats1Str} / (2軍)${stats2Str}`;
        const salaryStr = formatSalary(player['年俸']);
        const ageNum = String(player['年齢']).replace(/歳$/, '');

        allPlayersData.push({
            id: player['No'] || (player['背番号'] + '_' + originalName),
            name: originalName,
            team: teamName,
            position: position,
            age: ageNum,
            salaryStats: `${salaryStr}<br>${statsCombined}`,
            codeOrder: Object.keys(teamCodeMap).indexOf(code)
        });
    });

    const teamSelector = document.getElementById('teamSelector');
    teamSelector.innerHTML = '';
    
    let sortedTeams = Array.from(teamsSet).sort((a, b) => {
        const indexA = Object.values(teamCodeMap).indexOf(a);
        const indexB = Object.values(teamCodeMap).indexOf(b);
        return (indexA > -1 ? indexA : 99) - (indexB > -1 ? indexB : 99);
    });

    sortedTeams.forEach(team => {
        teamSelector.innerHTML += `<option value="${team}">${team}</option>`;
    });
}

document.getElementById('teamSelector').addEventListener('change', function(e) {
    currentTeam = e.target.value;
    protectedIds.clear(); 
    updateUI();
});

function updateUI() {
    const container = document.getElementById('playersContainer');
    container.innerHTML = '';
    
    if (!currentTeam || !isDataLoaded) return;

    const teamPlayers = allPlayersData.filter(p => p.team === currentTeam);
    const positions = ['投手', '捕手', '内野手', '外野手'];
    
    positions.forEach(pos => {
        const playersInPos = teamPlayers.filter(p => p.position === pos);
        if (playersInPos.length === 0) return;

        const titleDiv = document.createElement('div');
        titleDiv.className = 'category-title';
        
        const protectedInPos = playersInPos.filter(p => protectedIds.has(p.id)).length;
        titleDiv.innerHTML = `<span>${pos}</span> <span style="font-size:14px; color:#0056b3;">プロテクト ${protectedInPos}/${playersInPos.length}</span>`;
        container.appendChild(titleDiv);

        const wrapperDiv = document.createElement('div');
        wrapperDiv.className = 'player-grid-container';

        const gridDiv = document.createElement('div');
        gridDiv.className = 'player-grid';

        playersInPos.forEach(p => {
            const card = document.createElement('div');
            const isProtected = protectedIds.has(p.id);
            card.className = `player-card ${isProtected ? 'protected' : ''}`;
            
            const checkIcon = isProtected ? ' <span style="color:#d4af37; font-size:10px; margin-left:2px;">✔</span>' : '';

            card.innerHTML = `
                <div class="player-info">
                    <span class="name">${p.name}${checkIcon}</span>
                    <span class="age">${p.age}歳</span>
                </div>
                <div class="stats">${p.salaryStats}</div>
            `;

            card.onclick = () => toggleProtect(p.id);
            gridDiv.appendChild(card);
        });

        wrapperDiv.appendChild(gridDiv);
        container.appendChild(wrapperDiv);
    });

    document.getElementById('protectCount').innerText = protectedIds.size;
}

function toggleProtect(playerId) {
    if (protectedIds.has(playerId)) {
        protectedIds.delete(playerId);
    } else {
        if (protectedIds.size >= MAX_PROTECT) {
            alert('プロテクト枠の上限（28人）に達しています。');
            return;
        }
        protectedIds.add(playerId);
    }
    updateUI();
}
