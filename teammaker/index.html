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

        /* カードデザイン */
        .card { border: 1px solid #ddd; border-radius: 8px; padding: 20px; background: #fafafa; }
        .p-tag { display: inline-block; padding: 4px 12px; border-radius: 20px; background: #00796B; color: white; font-size: 12px; font-weight: bold; margin-bottom: 10px; }
        .p-name { font-size: 26px; font-weight: bold; margin-bottom: 5px; letter-spacing: 1px;}
        .p-meta { font-size: 14px; color: #666; margin-bottom: 15px; }

        /* パワプロ風 能力値スタイル */
        .stats-box { display: grid; gap: 8px; margin-top: 15px; margin-bottom: 15px; }
        .grid-4 { grid-template-columns: repeat(4, 1fr); }
        .grid-3 { grid-template-columns: repeat(3, 1fr); }
        .stat-card { background-color: #f2efe9; border: 1px solid #dcd9d2; border-radius: 6px; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 10px 0; box-shadow: 0 1px 2px rgba(0,0,0,0.05); }
        .stat-label { font-size: 11px; color: #555; font-weight: bold; margin-bottom: 4px; }
        .stat-rank { font-size: 28px; font-weight: 900; line-height: 1; margin-bottom: 4px; font-family: "Arial Black", "Impact", sans-serif; -webkit-text-stroke: 0.5px rgba(0,0,0,0.1); }
        .stat-val { font-size: 13px; color: #666; font-weight: bold; text-align: center; }

        /* タグ */
        .pitch-types, .defense-positions { display: flex; flex-wrap: wrap; gap: 6px; }
        .pos-tag, .pitch-tag { background-color: #e8f5e9; color: #1b5e20; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; border: 1px solid #c8e6c9; }
        
        /* テーブル・起用設定 */
        table { width: 100%; border-collapse: collapse; margin-top: 10px; font-size: 14px; }
        th, td { border: 1px solid #ccc; padding: 8px; text-align: center; }
        th { background: #eee; }
        select { padding: 4px; font-size: 14px; border-radius: 4px; }

        /* 結果画面のタブ */
        .tab-group { display: flex; margin-bottom: 15px; border-bottom: 2px solid #ddd; }
        .tab { padding: 10px 20px; cursor: pointer; font-weight: bold; color: #777; transition: 0.2s; }
        .tab.active { color: #1976D2; border-bottom: 3px solid #1976D2; }
        .tab-content { display: none; overflow-x: auto; }
        .tab-content.active { display: block; }
    </style>
</head>
<body>
    <div class="container">
        <!-- 1. ローディング画面 -->
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

        <!-- 3. 起用設定画面 -->
        <div id="screen-setup" class="screen">
            <h2 style="text-align:center; margin-bottom:5px;">起用設定</h2>
            <p style="text-align:center; color:#666; font-size:13px; margin-top:0;">ポジションや打順を割り当ててペナントへ！</p>
            
            <h3 style="border-bottom: 2px solid #00796B; color: #00796B;">野手起用（スタメン・打順）</h3>
            <div id="setup-fielders"></div>
            
            <h3 style="border-bottom: 2px solid #E53935; color: #E53935; margin-top:30px;">投手起用（ローテ・中継ぎ・抑え）</h3>
            <div id="setup-pitchers"></div>
            
            <button class="btn-primary" onclick="startPennant()">この設定でペナントレースを開始！</button>
        </div>

        <!-- 4. 結果・成績発表画面 -->
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

    <!-- ★強制キャッシュクリア用 -->
    <script src="script.js?v=20261013_3"></script>
</body>
</html>
