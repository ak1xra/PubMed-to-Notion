# PubMed → Notion 自動論文収集システム（GAS）

Google Apps Script（GAS）を用いて、PubMed から論文情報を取得し、Notion データベースへ自動保存するシステムです。  
毎日・定期的に最新論文を自動収集するパイプラインとして利用できます。

---

## 📌 機能概要

- Google スプレッドシートに記載した検索クエリを定期的に走査
- PubMed API（ESearch + EFetch）で論文情報を取得
- タイトル、PMID、Abstract、DOI、URL、出版日などを取得
- Notion Database へ自動的に Page を作成
- PMID をキーとして重複登録を防止（number プロパティで照合）
- GAS トリガー（時間ベース）で完全自動化

---

## 📁 ディレクトリ構成（推奨）

/
├─ src/
│   └─ main.gs
├─ README.md
├─ gas-pubmed-to-notion.js  ←←←←←←←←←← GAS
├─ docs/
│   ├─ notion-setup.md　(準備中)
│   ├─ gas-setup.md　(準備中)
│   └─ changelog.md　(準備中)
└─ LICENS

---

## 🧪 使用技術

- Google Apps Script (GAS)
- PubMed API（NCBI E-Utilities）
- Notion API v2022-06-28
- Google Spreadsheet

---

## 🗂 必要な準備

### 1. Notion Integration の作成
1. https://www.notion.com/my-integrations で Integration を作成
2. "Internal Integration Token" を取得  
3. 対象のデータベースへ Integration を共有

### 2. Notion データベース構造

| プロパティ名        | 型           | 説明                          |
|-------------------|--------------|------------------------------|
| Title             | Title        | 論文タイトル                 |
| PMID              | Number       | PubMed ID（重複チェックに利用） |
| PubMed URL        | URL          | PubMed ページへのリンク       |
| Publication Date  | Date         | 論文の公開日                 |
| DOI URL           | URL          | https://doi.org/...         |
| Abstract          | Rich Text    | 概要                         |

※ プロパティ名は GAS のコードと一致している必要があります。

---

## 3. スプレッドシートの準備

シート名：`Queries`

| A列 |
|----|
| Query（任意のヘッダー） |
| ASD |
| ADHD OR "attention-deficit hyperactivity disorder" |
| Genetics |
| "Genome-wide association study" |
| ... |

※ A2 以降に検索クエリを自由に追加

---

## 4. Script Properties の設定

**GAS → プロジェクト設定 → スクリプトプロパティ**

| キー名                | 値                                  |
|-----------------------|-------------------------------------|
| NOTION_API_KEY        | Notion インテグレーションの Token |
| NOTION_DATABASE_ID    | 対象 Notion DB の ID               |
| SPREADSHEET_ID        | クエリ管理スプシの ID              |

---

## ▶️ 実行方法

### 初回のみ
GAS エディタの関数選択から

fetchAllPubMedQueries

を手動実行し、OAuth 認証を通す。

### 自動化
GAS → トリガー → 新規作成  
時間主導型トリガーで **1日1回**などを設定。

---

## 📝 エラー例と対処

### ✔ `API token is invalid`
- NOTION_API_KEY を誤記（例：TIKEN など）
- Integration をデータベースに共有していない

### ✔ `PMID is expected to be number.`
- Notion 側の PMID が text 型 → number 型に変更
- GAS の送信が number になっていない → 修正版コードを適用

### ✔ `Queries シートが存在しません`
- シート名が `Queries` と一致していない

---

## 📜 ライセンス（任意）

MIT License など好みで設定。

---

## 📌 補足

本システムは「個人の研究支援ツール」として設計されています。  
論文管理の自動化・継続監視・要約生成（Notion AI などと組み合わせ）により、  
自分専用の研究ナレッジベースを構築できます。

---

## 💡 連絡・改善案

Issue または PR にて改善案を歓迎しています。


⸻
