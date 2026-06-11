# gas-scripts

Google Apps Script (GAS) プロジェクトのスクリプト管理リポジトリ。

## プロジェクト概要

Google Apps Script を使った自動化スクリプトをまとめて管理するリポジトリ。

## 技術スタック

- Google Apps Script (JavaScript)
- clasp (Google Apps Script CLI) ※必要に応じて使用

## Git 運用ルール

**コードを変更するたびに GitHub へプッシュすること。**

具体的な手順：

1. 変更をステージング
   ```
   git add <変更したファイル>
   ```

2. コミット（変更内容を端的に記述）
   ```
   git commit -m "変更内容の説明"
   ```

3. GitHub へプッシュ
   ```
   git push origin main
   ```

### コミットメッセージの規則

- 日本語・英語どちらでも可
- 変更の「目的」を書く（何を変えたかではなく、なぜ変えたか）
- 例: `スプレッドシートへの書き込みを日次で自動実行するよう変更`

### ブランチ戦略

- `main` ブランチを常に動作する状態に保つ
- 実験的な変更は feature ブランチを切って作業する

## 開発メモ

- GAS はブラウザ上のエディタまたは clasp でローカル編集できる
- clasp を使う場合は `clasp push` でデプロイする
