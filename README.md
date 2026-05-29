# 冷凍在庫システム

株式会社若林商店の冷凍庫在庫管理＆発注メール生成アプリ。

## URLs
- 在庫入力: https://ahonokiki.github.io/freezer-app/count.html
- 発注メール: https://ahonokiki.github.io/freezer-app/order.html

## 構成
- `count.html` - 小柏さん用カウント入力（PWA対応・iPhone想定）
- `order.html` - 賢さん用発注SMS生成（業者別タブ・自動入力）
- `index.html` - ランディングページ

データは Google Apps Script Webhook 経由でスプレッドシートに保存。
