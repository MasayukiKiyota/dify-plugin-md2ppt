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
| `config_yaml` | string | – | 自動判定の上書き。**通常は不要**（後述） |
| `file_name` | string | – | 出力ファイル名。既定は `presentation.pptx` |

`markdown_text` と `markdown_file` はどちらか一方を指定してください（両方指定した場合は
ファイルが使われ、その旨が警告に出ます）。

テンプレートのレイアウトは自動判定されるため、`config_yaml` を指定しなくても
日本語・英語・独自命名のどのテンプレートでもそのまま動きます。

出力は「テキストの要約 → JSON のメタデータ → `.pptx` ファイル」の順に返ります。
JSON には `slide_count` / `spec_count` / `layouts_used` / `outline` / `warnings` などが
入ります。

### 2. PowerPoint テンプレート解析 (`inspect_template`)

テンプレートのスライドサイズ・スライドマスタ・全レイアウトとそのプレースホルダを
一覧表示します。**自社テンプレートを使う前にまずこれを実行してください。**

| パラメータ | 型 | 必須 | 説明 |
| --- | --- | --- | --- |
| `template_file` | file | **必須** | 解析する `.pptx` / `.potx` |
| `emit_config` | boolean | – | `true` で `config_yaml` の雛形も生成し、`config.yaml` として返します（自動判定の内容を確認・調整したいとき用） |

出力例:

```
[1] 'タイトルとコンテンツ'
      idx  type                 name                      left   top    width  height  (inch)
      0    TITLE                Title 1                     0.92   0.49  11.50   1.25
      1    OBJECT               Content Placeholder 2       0.92   1.75  11.50   4.94
```

## 自社テンプレートを使う手順

**通常は設定不要です。** テンプレートをアップロードするだけで、使うべきレイアウトが
自動判定されます。判定は次の 2 段階です。

1. **レイアウト名のキーワード** — 「表紙 / 章扉 / 本文 / タイトルのみ / 白紙」や
   「Title Slide / Section Header / Title and Content / Title Only / Blank」など
2. **プレースホルダの構成** — 名前が手がかりにならない場合の判定基準

   | 用途 | 判定基準 |
   | --- | --- |
   | 表紙 | サブタイトルのプレースホルダを持つ（無ければ中央タイトル） |
   | 本文 | タイトル＋本文プレースホルダがちょうど 1 つ（2 カラムや比較は避ける） |
   | タイトルのみ | タイトルはあるが本文プレースホルダが無い |
   | 白紙 | プレースホルダが（日付・フッター・ページ番号を除いて）無い |
   | 章扉 | 「タイトルのみ」→「本文」の順に代用 |

   `placeholders` の `title` / `subtitle` / `body` の idx も実際のテンプレートから読み取ります。

実際に使われたレイアウト名は、変換結果の JSON の `layouts_used` で確認できます。

### 自動判定を上書きしたいとき

判定が意図と違う場合だけ `config_yaml` を指定してください。**明示した項目が常に優先され、
書かなかった項目は自動判定の値が使われます**（部分的な指定で構いません）。

```yaml
layouts:
  content: "タイトルとコンテンツ"   # ここだけ上書き。他は自動判定のまま
```

雛形が欲しい場合は `inspect_template` に `emit_config: true` を渡すと、そのテンプレート用の
設定が生成されます。

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

指定したレイアウト名がテンプレートに存在しない場合は、代わりに先頭のレイアウトが使われ、
`warnings` にその旨が出ます。

> **全角文字に注意**
> `config_yaml` に**全角コロン「：」**や**全角スペース**が混ざっていると YAML として
> 正しく読めません。全角コロンは構文エラーにすらならず、設定全体がただの文字列として
> 扱われます。IME を切り替えて入力した際に混入しやすいので、うまく反映されないときは
> まずここを疑ってください。エラーメッセージでも具体的に指摘します。

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
- **既定はテンプレート任せ（フォント・文字サイズ・配色を上書きしない）**
  `fonts` / `sizes` / `colors` / `table.style_id` / `spacing.list_indent` の
  既定値はすべて `null`（＝指定しない）です。設定なしで変換すると書体・文字
  サイズ・色・表スタイル・字下げを一切書き込まないので、スライドマスターと
  テーマの設定がそのまま残り、テンプレートの見た目を崩しません。
  空文字 `""` も `null` と同じ扱いです。
- **文字サイズもテンプレートに合わせる**
  `sizes` を `null` にすると `sz` を書き込まないので、テンプレートの文字サイズが
  そのまま効きます。あふれ判定に必要な実寸は、テンプレートから読み取ります
  （レイアウトのプレースホルダ → スライドマスターの `txStyles` →
  `presentation.xml` の既定）。本文プレースホルダに流すスライドと、表やコードを
  自前配置するスライドでは PowerPoint が適用する継承元が違うので、
  それぞれの継承元を見て見積もります。例外が 2 つあります。
  小見出し（`###` / `####`）はテンプレートに対応する書式が無いため、本文サイズ
  から比率で算出して書き込みます（書かないと本文と同じ大きさになるため）。
  また 1 枚に収まらず自動縮小するときは、縮小したサイズを書き込みます
  （`sz` を書かずに縮めることが PowerPoint の形式上できないため）。
- **色や書体を固定したいとき**
  `config_yaml` で値を書けば、そのぶんだけテンプレートより優先されます。
  フォントは `<a:latin>` だけでなく `<a:ea>` も明示するので、日本語が別フォントに
  落ちません（`fonts.eastasian` で指定）。表のセルまで塗って環境差をなくしたい
  場合は `colors.table_*` を書いたうえで `table.explicit_format: true` にします
  （既定は `false` ＝ 表スタイル側の書式に任せる）。
- **校正言語もテンプレートに合わせる**
  テンプレートに書かれている校正言語を読み取り、スライド・ノート・
  レイアウト・マスターの全文に書き込みます（`options.language: auto`）。
  日本語のテンプレートを使っても生成物が英語扱いになる問題を避けられます。
  あわせて「スペル チェックと文章校正を行わない」をオンにします（`options.no_proof: true`）。
  PowerPoint の「校閲 > 言語 > 校正言語の設定」で確認できます。
- **あふれの自動処理**
  1 枚に収まらない場合、まずフォントを段階的に縮小し（`options.shrink_steps`）、
  それでも収まらなければ「（続き）」スライドに分割します。表は行単位で分割し、
  ヘッダ行を各ページで繰り返します。このため `slide_count` は `spec_count`
  （見出しから決まる枚数）より多くなることがあります。

## `config_yaml` の主な設定項目

`layouts` / `placeholders`（自動判定される）のほかに、以下が指定できます。
指定しなかった項目は既定値です。`fonts` / `sizes` / `colors` / `table.style_id` /
`spacing.list_indent` の既定値は `null`（テンプレート任せ）なので、下の例は
「テンプレートより優先して固定したいとき」の書き方です。

```yaml
fonts:                       # 既定は 4 つとも null（テーマのフォントを使う）
  latin: Calibri
  eastasian: Yu Gothic       # Meiryo / MS PGothic / Noto Sans JP なども可
sizes:                       # 既定は 15 項目すべて null（テンプレートの大きさ）
  title: 32
  body: [18, 16, 14, 13, 12] # 箇条書きのレベル別。要素ごとに null も可
  min_body: 10               # 自動縮小の下限。null なら 10pt
colors:                      # 既定は 14 色すべて null（テーマの配色を使う）
  accent: "1E6F8E"
table:
  # 既定は style_id: null（テンプレートの表スタイル）。固定したいときだけ書く
  style_id: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"   # Medium Style 2 - Accent 1
  explicit_format: true      # 既定は false。true でも null の色は塗らない
spacing:
  line_ratio: 1.38           # 大きくすると早めに分割される
  list_indent: 0.3           # 既定は null（テンプレートのインデントを継承）
options:
  language: auto             # テンプレートの校正言語を複写。"ja-JP" で固定、null で書かない
  no_proof: true             # スペルチェックと文章校正を行わない
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
- 校正言語は**生成した .pptx のスライドマスター・レイアウトにも書き込まれます**
  （アップロードしたテンプレートのファイル自体は変更しません）。レイアウトごとに
  言語を分けている多言語テンプレートでは、その区別が失われます。避けたい場合は
  `options.language: null` と `options.no_proof: null` にしてください。
- 校正言語の自動判定はテンプレートの `lang` 属性の最頻値です。テンプレートが英語で
  作られている場合は英語と判定されるので、`options.language: ja-JP` で固定してください。
- `sizes` を `null`（既定）にすると、同じ内容でも**スライドの種類によって文字の
  大きさが変わることがあります**。テキストだけのスライドは本文プレースホルダの
  大きさを、表やコードを含むスライドはテンプレートの既定テキストの大きさを
  継承するためです。揃えたい場合は `sizes` に数値を書いてください
  （`sample/config.yaml` が全項目を書いた見本です）。
- `sizes` が `null` のときは、インラインコードを周囲より 6% 小さくする調整が
  効きません（`sz` を書かずに相対的に縮める手段が PowerPoint の形式に無いため）。
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
