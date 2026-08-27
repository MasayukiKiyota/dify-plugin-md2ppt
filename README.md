# Markdown to PowerPoint

**Author:** masayukikiyota
**Version:** 0.0.1
**Type:** Tool

Markdown を、既存の PowerPoint テンプレート（`.pptx` / `.potx`）のレイアウト・テーマ配色・
フォントに従って `.pptx` に変換する Dify プラグインです。見出し・表・箇条書き・コード・
引用・画像を、崩さずにスライド化することを目的にしています。

LLM が生成した Markdown をそのまま渡せば、自社テンプレートに沿った提案書・報告書が
ワークフローの出力としてダウンロードできます。

## ツール

### 1. Markdown を PowerPoint に変換 (`md_to_pptx`)

Markdown とテンプレートを受け取り、`.pptx` ファイルを返します。

| パラメータ | 型 | 必須 | 説明 |
| --- | --- | --- | --- |
| `markdown_text` | string | – | 変換する Markdown 本文。LLM ノードの出力を直結できます |
| `markdown_file` | file | – | 変換する `.md` / `.txt` ファイル。`markdown_text` より優先されます |
| `template_file` | file | **必須** | レイアウトとテーマを流用する `.pptx` / `.potx` |
| `config_yaml` | string | – | Markdown 要素とテンプレートのレイアウト名の対応づけ（後述） |
| `file_name` | string | – | 出力ファイル名。既定は `presentation.pptx` |

`markdown_text` と `markdown_file` はどちらか一方を指定してください（両方指定した場合は
ファイルが使われ、その旨が警告に出ます）。

出力は「テキストの要約 → JSON のメタデータ → `.pptx` ファイル」の順に返ります。
JSON には `slide_count` / `spec_count` / `outline` / `warnings` などが入ります。

### 2. PowerPoint テンプレート解析 (`inspect_template`)

テンプレートのスライドサイズ・スライドマスタ・全レイアウトとそのプレースホルダを
一覧表示します。**自社テンプレートを使う前にまずこれを実行してください。**

| パラメータ | 型 | 必須 | 説明 |
| --- | --- | --- | --- |
| `template_file` | file | **必須** | 解析する `.pptx` / `.potx` |
| `emit_config` | boolean | – | `true` で `config_yaml` の雛形も生成し、`config.yaml` として返します |

出力例:

```
[1] 'タイトルとコンテンツ'
      idx  type                 name                      left   top    width  height  (inch)
      0    TITLE                Title 1                     0.92   0.49  11.50   1.25
      1    OBJECT               Content Placeholder 2       0.92   1.75  11.50   4.94
```

## 自社テンプレートを使う手順

1. `inspect_template` に自社テンプレートを渡し、`emit_config` を `true` にする
2. 返ってきた `config_yaml`（または `config.yaml` ファイルの中身）を確認する
3. それを `md_to_pptx` の `config_yaml` パラメータに貼り付ける

レイアウト名は「表紙 / 章扉 / 本文 / タイトルのみ / 白紙」といったキーワードから推測されます。
推測が外れている場合は手で直してください。

```yaml
layouts:
  title:   "タイトル スライド"   # front matter の title から作る表紙
  section: "セクション見出し"     # H1
  content: "タイトルとコンテンツ" # H2（テキストのみのスライド）
  table:   "タイトルのみ"         # 表・コード・画像を含むスライド
  blank:   "白紙"
placeholders:
  title: 0
  subtitle: 1
  body: 1
```

`config_yaml` を指定しない場合は既定値が使われます。テンプレートのレイアウト名が
既定値と一致しないときは、そのレイアウトの代わりに先頭のレイアウトが使われ、
`warnings` にその旨が出ます。**警告が出たら `config_yaml` を設定してください。**

## Markdown の対応表

| Markdown | スライドでの扱い |
| --- | --- |
| front matter (`title` / `subtitle` / `author` / `date`) | 表紙スライド |
| `#` | **章扉スライド**（`options.section_slides: false` で無効化） |
| `##` | **新規スライド**（タイトルになる） |
| `###` / `####` | スライド内の小見出し（アクセント色・太字） |
| 段落 | 本文テキスト |
| `-` / `*` リスト | 箇条書き（ネスト 5 階層まで、レベル別フォントサイズ） |
| `1.` リスト | 番号付きリスト（PowerPoint の自動採番） |
| 表 | **ネイティブの PowerPoint 表**（列幅を内容から自動配分、`---:` `:---:` の寄せも反映） |
| ` ``` ` コードブロック | 等幅フォント＋背景付きの角丸ボックス |
| `>` 引用 | 斜体＋縦罫線付きのブロック |
| `![alt](path)` | 画像（後述の制限あり） |
| `---` | 区切り線 |
| `**太字**` `*斜体*` `` `コード` `` `~~打消~~` `[リンク](url)` | 文字単位で書式を反映 |
| `<!-- notes: ... -->` | スピーカーノート |

## 資料として崩れないための仕組み

- **テキストだけのスライドは本文プレースホルダに流し込む**
  テンプレートの箇条書き記号・行間・自動縮小（autofit）がそのまま効きます。
- **表・コード・画像を含むスライドは手動レイアウト**
  「タイトルのみ」レイアウトを使い、本文プレースホルダと同じ領域にブロックを縦に積みます。
  位置は `body_area` で上書きできます。
- **日本語フォントの指定**
  `<a:latin>` だけでなく `<a:ea>` も明示するため、日本語が別フォントに落ちません
  （`fonts.eastasian` で指定）。
- **あふれの自動処理**
  1 枚に収まらない場合、まずフォントを段階的に縮小し（`options.shrink_steps`）、
  それでも収まらなければ「（続き）」スライドに分割します。表は行単位で分割し、
  ヘッダ行を各ページで繰り返します。このため `slide_count` は `spec_count`
  （見出しから決まる枚数）より多くなることがあります。

## `config_yaml` の主な設定項目

`layouts` / `placeholders` のほかに、以下が指定できます。指定しなかった項目は既定値です。

```yaml
fonts:
  latin: Calibri
  eastasian: Yu Gothic       # Meiryo / MS PGothic / Noto Sans JP なども可
sizes:
  title: 32
  body: [18, 16, 14, 13, 12] # 箇条書きのレベル別
colors:
  accent: "1E6F8E"
table:
  # テンプレート側の表スタイルに任せたいときは style_id: null / explicit_format: false
  style_id: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"   # Medium Style 2 - Accent 1
  explicit_format: true
spacing:
  line_ratio: 1.38           # 大きくすると早めに分割される
  list_indent: 0.3           # null でテンプレートのインデントを継承
options:
  auto_split: true
  shrink_steps: 2            # 分割前に何段階フォントを縮めるか（1 段 = 8%）
  section_slides: true       # false にすると H1 も通常スライドに
  page_number: true
```

よく使う組み込み表スタイルの GUID:

| スタイル名 | GUID |
| --- | --- |
| No Style, Table Grid | `{5940675A-B579-460E-94D1-54222C63F5DA}` |
| Light Style 1 – Accent 1 | `{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}` |
| Medium Style 2 – Accent 1 | `{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}` |
| Medium Style 4 – Accent 1 | `{22838BEF-8BB2-4498-84A7-C5851F593DF1}` |
| Dark Style 1 – Accent 1 | `{B301B821-A1FF-4177-AEE7-76D212191A09}` |

## 制約・既知の注意点

- **画像は現状ほぼ使えません。** `![alt](path)` の相対パスは変換用の一時ディレクトリを
  基準に解決されますが、そこには何も置かれないため見つかりません。`http(s)://` の画像も
  取得しません（意図的にネットワークアクセスを行わないため）。いずれの場合も
  「画像が見つかりません」というプレースホルダが描かれ、`warnings` に記録されます。
- SVG / EMF 画像は python-pptx が読めないため、PNG などに変換してから使ってください。
- テンプレートのレイアウトを**新規に作ることはできません**（python-pptx の制約）。
  必要なレイアウトはあらかじめ PowerPoint 側で用意してください。
- 段組みレイアウト（2 カラムなど）への自動振り分けは未対応です。
- あふれ判定は文字幅の推定に基づく近似です。行送りが極端なテンプレートでは
  `spacing.line_ratio` を調整してください。
- スピーカーノート（`<!-- notes: ... -->`）は、変換元の CLI と同じく最初の 1 件だけが
  先頭スライドに付きます。
- 古い `.ppt` 形式は非対応です。PowerPoint で `.pptx` に保存し直してください。

## 開発

```bash
uv venv
uv pip install -r requirements.txt

# 変換パイプラインの検証（dify_plugin 不要）
.venv/Scripts/python.exe tests/test_convert.py
# ツールの _invoke を実際に回す検証（dify_plugin 必要）
.venv/Scripts/python.exe tests/test_tools.py
```

リモートデバッグ:

```bash
cp .env.example .env     # Dify の「プラグイン → デバッグ」で取得したキーを設定
.venv/Scripts/python.exe -m main
```

パッケージング:

```bash
dify plugin package .
```

### 変換ロジックの取り込み方

[md2pptx_core.py](md2pptx_core.py) は `sample/md2pptx.py` の**バイト同一のコピー**です。
Dify 固有の事情（ファイル blob の入出力、一時ファイル、`SystemExit`、`.potx` の
content type）はすべて [md2ppt_utils.py](md2ppt_utils.py) 側で吸収しているため、
`md2pptx_core.py` は決して編集しないでください。上流が更新されたらコピーし直すだけで
追随できます。

```bash
cp sample/md2pptx.py md2pptx_core.py && diff -q sample/md2pptx.py md2pptx_core.py
cp sample/sample.md sample/template.pptx sample/config.yaml tests/fixtures/
```

`sample/` は上流の取り込み元であり、`.gitignore` で除外されています（ローカルにのみ
存在します）。リポジトリをクローンしただけの環境でもテストが回るよう、テストが使う
3 ファイルは [tests/fixtures/](tests/fixtures/) に複製してあります。

## Requirements

- Python 3.12
- dify_plugin >= 0.9.0
- python-pptx / markdown-it-py / PyYAML / Pillow
