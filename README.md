# Patient-information（患者説明文書リポジトリ）

本リポジトリは、耳鼻咽喉科・頭頸部外科を中心とした診療現場で、患者さんとご家族への説明を支援するための患者説明文書を管理する場所です。

Markdown 原本から Word・PDF の患者説明文書を作成し、確認済みのアイコンやイラストを添えたうえで、説明用スライドと試作・閲覧用HTMLへ展開します。医療内容と図版の最終確認は、必ず医師が行います。

## 概要

- **目的**: 疾患説明、検査説明、手術説明、治療方針説明の品質を平準化し、患者さんに伝わりやすい説明資料を継続的に整備する。
- **主な対象**: 耳科、鼻副鼻腔、咽喉頭、頸部、甲状腺、麻酔関連などの患者向け説明文書。
- **管理言語**: 主に日本語。一部、多言語対応の文書も含む。
- **正本**: `src/*.md` の Markdown ファイル。
- **生成物**: `docx/*.docx` のWord文書、`PDF/*.pdf` の配布用PDF、`pptx/*.pptx` の説明用スライド、`HTML/*.html` の試作・閲覧用HTML。

## ディレクトリ構成

```text
Patient-information/
├── src/                 # 患者説明文書の Markdown 正本
├── src_202605原本/      # 2026年5月時点の参照用原本。改変しない
├── images/              # アイコン・イラスト（master: 原版、optimized: 文書用）
├── convert_to_office.py # Markdown から Word (.docx) & PDF を生成するスクリプト
├── convert_to_pptx.py   # Markdown から説明用スライド (.pptx) を生成するスクリプト
├── generate_html.py     # Markdown から試作・閲覧用HTMLを生成するスクリプト
├── HTML/                # 試作・閲覧用HTMLの生成先
├── README.md            # この説明書
├── CHANGELOG.md         # 変更履歴
└── .gitignore           # 開発用の除外設定
```

関連する根拠メモは、Obsidian Vault 側の以下に保存します。

```text
C:\Users\yert1\Documents\agy\ObsidianVault\30_Notes\Patient-information-sources\
```

## 運用方針

### 1. Markdown を唯一の正本にする

医学的な本文は `src/*.md` を正本として管理します。文書を修正する場合は、必ず Markdown を更新してから Word・PDF・PowerPoint・HTMLを再生成します。HTMLを直接編集して医学的内容を更新してはいけません。

`src_202605原本/` は、今回の改訂プロジェクトを始める前の原本を残すための参照用ディレクトリです。改訂時に内容を確認しても、このディレクトリ内のファイルは変更しません。Obsidian の根拠メモを反映した改訂版は `src/` に作成します。

### 2. Markdownから文書、図版付き文書、スライドへ展開する

`src/*.md` には、病態、治療目的、術式、代替治療、合併症、術後注意、再発、長期通院、同意確認に必要な情報をできるだけ保持します。Markdown を短くしすぎると、医師レビュー、同意書、印刷、将来の再利用に必要な情報が失われるためです。

まず Markdown から医師レビュー用のWord文書と配布用PDFを作成します。図版を使用する場合は、医師が内容を確認したアイコンまたはイラストを `images/optimized/` に置き、本文中で次のように指定します。

```markdown
![鼓膜の位置を示す図](../images/optimized/鼓膜の位置.png)
```

文書とスライドの双方に同じ図版を使えます。スライドでは「1スライド1メッセージ」を優先するため、必要なセクションに短い要約を付けます。

```markdown
{{slide_summary: 真珠腫を取り除き、進行を止めることが手術の主な目的です。}}
```

図版の原版は `images/master/`、文書用に最適化した版は `images/optimized/` に保存します。各図版には、用途、作成指示、医師確認状況を記録します。

### 3. Word・PDF・スライド・HTMLを同一世代で管理する

Word・PDF・スライド・HTMLは Markdown 正本から生成する成果物です。HTMLは、患者さんが画面で読む際の見やすさや、アイコン・イラスト・注意事項の配置を確認するための試作・閲覧版です。原稿や図版を更新した場合は、対応する成果物を再生成し、同じ版として確認します。

つまり、文書を更新した場合は原則として以下を同じ変更単位で扱います。

- `src/*.md`
- `images/optimized/` の確認済み図版（使用時）
- 対応する `docx/*.docx`、`PDF/*.pdf`、`pptx/*.pptx`
- 対応する `HTML/*.html`（画面表示を確認する場合）
- 必要に応じて `CHANGELOG.md`

### 4. Obsidian は根拠メモと運用ルールの管理場所

Obsidian Vault には、NotebookLM 由来の要約、文献メモ、改訂根拠、運用ガイドラインを置きます。患者説明文書を改訂する場合は、必要に応じて `Patient-information-sources` の根拠メモを参照します。

### 5. 医師レビューを必須にする

AI による文書整形や表現調整は補助作業です。臨床で使用する前に、医学的妥当性、合併症、禁忌、受診目安、患者さんへの伝わり方を医師が確認します。

未確認の文書は、フロントマターで `status: "draft"`、`reviewed_by: "未確認"` のように明示します。

### 6. 個人情報を入れない

本リポジトリには、特定の患者さんを識別できる情報、実患者データ、個人健康情報（PHI）を含めません。説明文書は汎用的な患者向け資料として管理します。

## Markdown 正本の基本形式

新規作成または大幅改訂時は、可能な限り以下のフロントマターを維持します。

```yaml
---
title: "文書タイトル"
version: "0.1.0"
status: "draft"
reviewed_by: "未確認"
review_date: ""
evidence_source: "根拠メモまたは文献名"
change_reason: "今回の改訂理由"
updated: "YYYY-MM-DD"
---
```

## 各種ファイル生成手順

### 1. Office ドキュメント（Word/PDF）生成
Markdown 正本から Word 形式（`.docx`）および閲覧用 PDF（`.pdf`）を生成するには、以下を実行します。

```bash
python convert_to_office.py src/対象ファイル名.md
```

実行後、`docx/` および `PDF/` 配下に対応するファイルが出力されます。DOCX生成後には、`C:\Users\yert1\Documents\agy\00_System\bin\officecli.exe` を用いてOpenXMLスキーマ検証と文書問題の検出を行います。検証結果は表示しますが、警告だけではPDF化を中断しません。

### 2. 説明用スライド（PowerPoint）生成
Markdown 正本から説明用スライド（`.pptx`）を生成するには、以下を実行します。スライド要約または本文抜粋がスライドに反映され、スピーカーノートに Markdown の詳細テキストが自動流し込みされます。Markdownに指定した図版は対応するスライドにも配置されます。

```bash
python convert_to_pptx.py src/対象ファイル名.md
```

実行後、`pptx/` 配下に対応する `*_説明スライド.pptx` ファイルが出力されます。生成後にはOfficeCLIでOpenXMLスキーマ検証とスライド問題の検出を行い、警告だけでは出力処理を中断しません。

### 3. 試作・閲覧用HTML生成

HTMLは、Word・PDFの代替となる正式な診療文書ではありません。アイコン・イラスト、注意事項、文字サイズ、画面上の読みやすさを確認するために使います。

```bash
python generate_html.py src/対象ファイル名.md
```

実行後、`HTML/` 配下に同名の `.html` ファイルが出力されます。HTMLとWord・PPTXは同じMarkdown画像記法を使うため、確認済みの図版が共通して反映されます。

## 改訂時の基本フロー

1. Obsidian 側の根拠メモ、既存 Markdown、必要なガイドラインを確認する。
2. `src/*.md` の本文とフロントマターを更新する。
3. 禁忌、合併症、受診目安、フォロー方針の抜けがないか確認する。
4. `python convert_to_office.py src/対象ファイル名.md` でWord・PDFを生成する。
5. Word文書に図版を反映し、内容・視認性・図版の医学的妥当性を確認する。
6. `python convert_to_pptx.py src/対象ファイル名.md` で説明用スライドを生成し、要約・図版・発表者ノートを確認する。
7. `python generate_html.py src/対象ファイル名.md` でHTMLを生成し、画面での読みやすさと図版配置を確認する。
8. 医師レビュー後、`reviewed_by`、`review_date`、`status` を更新する。
9. Markdown、図版、Word・PDF・PowerPoint・HTML、変更履歴を同じ世代として Git に記録する。

## 注意事項

- HTMLは試作・閲覧版です。医学的内容の変更は必ずMarkdown正本へ戻して行ってください。
- 図版は、患者さんに不安を与えにくい簡潔な表現とし、使用前に解剖学的・医学的な妥当性を医師が確認してください。
- 医療者向けの専門的な手順書や内部資料は、患者向け説明文書と混在させないでください。
- PDF や画像を本格的に扱う場合は、Obsidian 側の運用ガイドラインに従い、保存場所と Git 管理方針を確認してください。
- 薬価、助成制度、診療ガイドラインなど変わり得る情報は、改訂時に最新情報を確認してください。
