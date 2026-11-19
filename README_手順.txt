ドリームベースボールバトル Supabase版 v0.1

1) Supabase SQL Editorで setup.sql を実行。
2) index.htmlをGitHub Pagesへ置く。
3) サイトを開き、Supabase Publishable keyを接続設定欄へ入力→保存→接続テスト。
4) 2台で同じURLを開き、ルーム作成/参加テスト。

Project URLは設定済み: https://jrvwjrbjhknkylvkcocn.supabase.co
Publishable keyはSupabase DashboardのProject Settings > API Keysから取得。service_role/secret keyは絶対にブラウザへ入力しない。
この試作はRLSが全許可のため、知っている人同士のテスト専用。公開前にルーム単位の権限制限を実装する。
