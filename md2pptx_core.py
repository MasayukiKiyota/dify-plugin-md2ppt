#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
md2pptx — Markdown を PowerPoint テンプレートに流し込んで .pptx を生成する。

  python3 md2pptx.py input.md -o output.pptx -c config.yaml

依存: python-pptx, markdown-it-py, PyYAML, Pillow
  pip install python-pptx markdown-it-py PyYAML Pillow

見出しの対応（既定）:
  #   -> 章扉スライド
  ##  -> 新規スライド（タイトル）
  ### -> スライド内の小見出し
  #### 以下 -> 小見出し（やや小さめ）

対応要素: 段落 / 箇条書き・番号付きリスト（多階層）/ 表 / コードブロック /
          引用 / 画像 / 水平線 / 太字・斜体・インラインコード・リンク
"""
from __future__ import annotations

import argparse
import copy
import math
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Sequence

import yaml
from markdown_it import MarkdownIt
from markdown_it.tree import SyntaxTreeNode
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

EMU_PER_IN = 914400
EMU_PER_PT = 12700


# ════════════════════════════════════════════════════════════════════════
# 1. 設定
# ════════════════════════════════════════════════════════════════════════

DEFAULT_CONFIG: dict[str, Any] = {
    "template": "template.pptx",
    "layouts": {
        "title": "表紙",
        "section": "章扉",
        "content": "本文",
        "table": "タイトルのみ",   # 表・コード・画像を含むスライド
        "blank": "白紙",
    },
    "placeholders": {"title": 0, "subtitle": 1, "body": 1},
    # 手動レイアウト時の描画領域（inch）。null は本文プレースホルダから自動取得
    "body_area": {"left": None, "top": None, "width": None, "height": None},
    # null は「指定しない」。既定では 4 つとも null で、書体を一切書き込まず
    # スライドマスター（テーマ）のフォントをそのまま使う。
    "fonts": {
        "latin": None,
        "eastasian": None,
        "code": None,
        "code_eastasian": None,
    },
    # null は「指定しない」。sz を書き込まないので、テンプレート（レイアウトの
    # プレースホルダ → スライドマスター）の文字サイズがそのまま効く。高さの
    # 見積りにはテンプレートから読み取った実効サイズを使う。
    # 既定ではすべて null ＝ テンプレートの文字サイズをそのまま使う。
    "sizes": {
        "title": None,
        "section_title": None,
        "cover_title": None,
        "cover_subtitle": None,
        "h3": None,
        "h4": None,
        "body": None,                   # 箇条書きレベル別。[18, 16, ...] と書ける
        "paragraph": None,
        "table": None,
        "table_header": None,
        "code": None,
        "quote": None,
        "caption": None,
        "page_number": None,
        "min_body": None,               # 自動縮小の下限。null は組み込みの 10pt
    },
    # null は「指定しない」。その書式を書き込まないので、文字はテーマの色を
    # 継承し、図形・表はテーマ／表スタイルの既定書式のままになる。
    # 既定ではすべて null ＝ テンプレートの配色をそのまま使う。
    "colors": {
        "text": None,
        "heading": None,
        "accent": None,
        "muted": None,
        "code_text": None,
        "code_bg": None,
        "quote_text": None,
        "quote_bar": None,
        "table_header_bg": None,
        "table_header_text": None,
        "table_band_bg": None,
        "table_body_bg": None,          # 帯にならない行の地色
        "table_text": None,
        "table_border": None,
    },
    "table": {
        # PowerPoint 組み込みスタイル ID。既定の null はテンプレート側の表スタイル。
        # 例: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"（Medium Style 2 - Accent 1）
        "style_id": None,
        "first_row_header": True,
        "banding": True,
        # 既定の false はテンプレート／表スタイル側の書式に任せる。
        # true にすると colors.table_* で明示的に塗る（どの環境でも同じ見た目）。
        # ただし塗るのは色を書いた項目だけで、null の項目は塗らない。
        "explicit_format": False,
        "min_row_height": 0.32,   # inch
        "cell_margin": 0.06,      # inch
    },
    "quote": {"bar": True, "indent": 0.3, "italic": True},
    "image": {"max_height": 4.2, "align": "center"},
    "spacing": {
        "block_gap": 8,          # pt
        "para_space_after": 6,   # pt
        "line_ratio": 1.38,
        # 箇条書きのインデント（inch）。既定の null はテンプレートの設定を継承
        "list_indent": None,
        "list_hanging": None,
    },
    "options": {
        "auto_split": True,               # 収まらない場合に続きスライドを作る
        "continuation_suffix": "（続き）",
        "shrink_steps": 2,                # 分割の前に何段階まで縮小するか
        "cover_from_first_h1": False,     # 先頭の H1 を表紙にする
        "page_number": True,
        "page_number_skip_first": True,
        "section_slides": True,           # H1 で章扉を作る
        "notes_marker": "notes",          # <!-- notes: ... --> をスピーカーノートに
        "table_header_repeat": True,      # 表を分割したらヘッダを繰り返す
        # 校正言語（PowerPoint の「校閲 > 言語 > 校正言語の設定」）。
        #   "auto"  … テンプレートに書かれている言語を複写する（既定）
        #   null/"" … lang を一切書き込まない（従来どおり）
        #   "ja-JP" … その言語をドキュメント全体に適用する
        "language": "auto",
        # 「スペル チェックと文章校正を行わない」（noProof）。
        #   true（既定）… オンにする / false … 明示的にオフ / null … 触らない
        "no_proof": True,
    },
}


def deep_merge(base: dict, over: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(path: str | None) -> dict:
    if not path:
        return copy.deepcopy(DEFAULT_CONFIG)
    data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
    return deep_merge(DEFAULT_CONFIG, data)


# ════════════════════════════════════════════════════════════════════════
# 2. Markdown 解析
# ════════════════════════════════════════════════════════════════════════

@dataclass
class Run:
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False
    strike: bool = False
    link: str | None = None


@dataclass
class Heading:
    level: int
    runs: list[Run]


@dataclass
class Para:
    runs: list[Run]


@dataclass
class ListItem:
    level: int
    ordered: bool
    number: int
    runs: list[Run]


@dataclass
class ListBlock:
    items: list[ListItem] = field(default_factory=list)


@dataclass
class Table:
    header: list[list[Run]] = field(default_factory=list)
    rows: list[list[list[Run]]] = field(default_factory=list)
    aligns: list[str] = field(default_factory=list)


@dataclass
class Code:
    lang: str
    text: str


@dataclass
class Quote:
    blocks: list = field(default_factory=list)


@dataclass
class Image:
    src: str
    alt: str = ""


@dataclass
class Rule:
    pass


Block = Heading | Para | ListBlock | Table | Code | Quote | Image | Rule

FRONT_MATTER = re.compile(r"\A---\s*\n(.*?)\n---\s*\n", re.S)


def split_front_matter(src: str) -> tuple[dict, str]:
    m = FRONT_MATTER.match(src)
    if not m:
        return {}, src
    try:
        meta = yaml.safe_load(m.group(1)) or {}
    except Exception:
        meta = {}
    return (meta if isinstance(meta, dict) else {}), src[m.end():]


def _inline_runs(node: SyntaxTreeNode) -> list[Run]:
    """inline ノード配下を Run の並びに変換する。"""
    runs: list[Run] = []
    state = {"bold": 0, "italic": 0, "strike": 0, "link": None}

    def walk(children: Iterable[SyntaxTreeNode]) -> None:
        for ch in children:
            t = ch.type
            if t == "text":
                if ch.content:
                    runs.append(Run(ch.content, state["bold"] > 0, state["italic"] > 0,
                                    False, state["strike"] > 0, state["link"]))
            elif t == "code_inline":
                runs.append(Run(ch.content, state["bold"] > 0, state["italic"] > 0,
                                True, state["strike"] > 0, state["link"]))
            elif t == "strong":
                state["bold"] += 1
                walk(ch.children)
                state["bold"] -= 1
            elif t == "em":
                state["italic"] += 1
                walk(ch.children)
                state["italic"] -= 1
            elif t == "s":
                state["strike"] += 1
                walk(ch.children)
                state["strike"] -= 1
            elif t == "link":
                prev, state["link"] = state["link"], ch.attrs.get("href")
                walk(ch.children)
                state["link"] = prev
            elif t in ("softbreak", "hardbreak"):
                runs.append(Run("\n"))
            elif t == "image":
                alt = "".join(c.content for c in ch.children or [])
                runs.append(Run(f"[{alt or '画像'}]"))
            else:
                if ch.children:
                    walk(ch.children)
                elif ch.content:
                    runs.append(Run(ch.content))

    walk(node.children or [])
    return [r for r in runs if r.text]


def _para_is_image(node: SyntaxTreeNode) -> Image | None:
    """段落が画像 1 個だけなら Image ブロックとして扱う。"""
    inline = next((c for c in node.children if c.type == "inline"), None)
    if inline is None:
        return None
    kids = [c for c in (inline.children or []) if not (c.type == "text" and not c.content.strip())]
    if len(kids) == 1 and kids[0].type == "image":
        img = kids[0]
        alt = "".join(c.content for c in img.children or [])
        return Image(img.attrs.get("src", ""), alt)
    return None


NOTES_RE = re.compile(r"<!--\s*(?:notes?|ノート)\s*[:：]\s*(.*?)\s*-->", re.S | re.I)


def _cells_of_row(tr: SyntaxTreeNode) -> list[list[Run]]:
    cells = []
    for cell in tr.children:
        inline = next((c for c in cell.children if c.type == "inline"), None)
        cells.append(_inline_runs(inline) if inline is not None else [])
    return cells


def _aligns_of_row(tr: SyntaxTreeNode) -> list[str]:
    out = []
    for cell in tr.children:
        style = (cell.attrs.get("style") or "")
        if "center" in style:
            out.append("center")
        elif "right" in style:
            out.append("right")
        else:
            out.append("left")
    return out


def nodes_to_blocks(nodes: Sequence[SyntaxTreeNode], level: int = 0,
                    notes: list[str] | None = None) -> list[Block]:
    blocks: list[Block] = []
    for node in nodes:
        t = node.type
        if t == "heading":
            inline = next((c for c in node.children if c.type == "inline"), None)
            blocks.append(Heading(int(node.tag[1]), _inline_runs(inline) if inline else []))
        elif t == "paragraph":
            img = _para_is_image(node)
            if img is not None:
                blocks.append(img)
            else:
                inline = next((c for c in node.children if c.type == "inline"), None)
                runs = _inline_runs(inline) if inline else []
                if runs:
                    blocks.append(Para(runs))
        elif t in ("bullet_list", "ordered_list"):
            ordered = t == "ordered_list"
            lb = ListBlock()
            _collect_list(node, ordered, level, lb, notes)
            if lb.items:
                blocks.append(lb)
        elif t == "table":
            tb = Table()
            for section in node.children:
                for tr in section.children:
                    if section.type == "thead":
                        tb.header = _cells_of_row(tr)
                        tb.aligns = _aligns_of_row(tr)
                    else:
                        tb.rows.append(_cells_of_row(tr))
            if not tb.aligns and tb.rows:
                tb.aligns = ["left"] * len(tb.rows[0])
            blocks.append(tb)
        elif t in ("fence", "code_block"):
            blocks.append(Code(node.info.strip() if node.info else "", node.content.rstrip("\n")))
        elif t == "blockquote":
            blocks.append(Quote(nodes_to_blocks(node.children, level, notes)))
        elif t == "hr":
            blocks.append(Rule())
        elif t == "html_block":
            if notes is not None:
                for m in NOTES_RE.finditer(node.content):
                    notes.append(m.group(1).strip())
        else:
            if node.children:
                blocks.extend(nodes_to_blocks(node.children, level, notes))
    return blocks


def _collect_list(node: SyntaxTreeNode, ordered: bool, level: int,
                  out: ListBlock, notes: list[str] | None) -> None:
    number = int(node.attrs.get("start", 1)) if ordered else 1
    for item in node.children:
        if item.type != "list_item":
            continue
        first = True
        for child in item.children:
            if child.type == "paragraph":
                inline = next((c for c in child.children if c.type == "inline"), None)
                runs = _inline_runs(inline) if inline else []
                if runs:
                    out.items.append(ListItem(level, ordered, number if first else 0, runs))
                    first = False
            elif child.type in ("bullet_list", "ordered_list"):
                _collect_list(child, child.type == "ordered_list", level + 1, out, notes)
            else:
                for b in nodes_to_blocks([child], level, notes):
                    if isinstance(b, Para):
                        out.items.append(ListItem(level, ordered, 0, b.runs))
        number += 1


def parse_markdown(src: str) -> tuple[dict, list[Block], list[str]]:
    meta, body = split_front_matter(src)
    md = MarkdownIt("commonmark").enable(["table", "strikethrough"])
    tokens = md.parse(body)
    tree = SyntaxTreeNode(tokens)
    notes: list[str] = []
    blocks = nodes_to_blocks(tree.children, 0, notes)
    return meta, blocks, notes


# ════════════════════════════════════════════════════════════════════════
# 3. スライド分割
# ════════════════════════════════════════════════════════════════════════

@dataclass
class SlideSpec:
    kind: str                 # "cover" | "section" | "content"
    title: list[Run] = field(default_factory=list)
    subtitle: list[Run] = field(default_factory=list)
    blocks: list[Block] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def runs_text(runs: Sequence[Run]) -> str:
    return "".join(r.text for r in runs)


def build_slides(meta: dict, blocks: list[Block], cfg: dict) -> list[SlideSpec]:
    opts = cfg["options"]
    slides: list[SlideSpec] = []
    cur: SlideSpec | None = None

    if meta.get("title"):
        sub = meta.get("subtitle") or meta.get("date") or meta.get("author") or ""
        extra = [meta.get(k) for k in ("author", "date") if meta.get(k) and meta.get(k) != sub]
        subtitle = " / ".join([s for s in [str(sub)] + [str(e) for e in extra] if s])
        slides.append(SlideSpec("cover", [Run(str(meta["title"]))],
                                [Run(subtitle)] if subtitle else []))

    def flush() -> None:
        nonlocal cur
        if cur is not None:
            slides.append(cur)
            cur = None

    pending_notes: list[str] = []

    for blk in blocks:
        if isinstance(blk, Heading) and blk.level == 1:
            flush()
            if not slides and opts["cover_from_first_h1"]:
                slides.append(SlideSpec("cover", blk.runs))
                continue
            if opts["section_slides"]:
                slides.append(SlideSpec("section", blk.runs))
            else:
                cur = SlideSpec("content", blk.runs)
            continue
        if isinstance(blk, Heading) and blk.level == 2:
            flush()
            cur = SlideSpec("content", blk.runs)
            continue
        if cur is None:
            cur = SlideSpec("content", [])
        cur.blocks.append(blk)

    flush()
    return [s for s in slides if s.kind != "content" or s.title or s.blocks]


# ════════════════════════════════════════════════════════════════════════
# 4. 描画ユーティリティ
# ════════════════════════════════════════════════════════════════════════

def char_units(s: str) -> float:
    """全角=1.0, 半角=0.52 として文字幅を数える簡易メトリクス。"""
    total = 0.0
    for ch in s:
        o = ord(ch)
        if o < 0x2E80 or 0xFF61 <= o <= 0xFF9F:
            total += 0.52
        else:
            total += 1.0
    return total


def wrapped_lines(text: str, size_pt: float, width_pt: float) -> int:
    if width_pt <= 0:
        return 1
    size_pt = max(size_pt, 1.0)             # 0 除算よけ（SizeBook が保証するが二重に）
    cap = max(width_pt / size_pt, 1.0)
    lines = 0
    for logical in text.split("\n"):
        u = char_units(logical)
        lines += max(1, math.ceil(u / cap))
    return max(lines, 1)


def rgb(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr.replace("#", "").upper())


def opt_str(value: Any) -> str | None:
    """null / 空文字を「指定なし」として None に揃える。

    fonts / colors の各項目は null にできる。None になった項目は書式を
    書き込まないので、テンプレート側の設定がそのまま残る。
    """
    return value if isinstance(value, str) and value.strip() else None


def opt_bool(value: Any) -> bool | None:
    """true / false / null を揃える。null と空文字は「指定なし」。

    Dify から真偽値が "true" / "false" という文字列で届くことがあるので、
    文字列も解釈する（bool_param と同じ考え方）。
    """
    if value is None:
        return None
    if isinstance(value, str):
        text = value.strip().lower()
        if not text:
            return None
        return text in ("1", "true", "yes", "on")
    return bool(value)


# OOXML の ST_TextFontSize は 100〜40000（1〜400pt）。Pt() もこの外だと落ちる。
MIN_FONT_SIZE = 1.0
MAX_FONT_SIZE = 400.0
# 自動縮小の下限（sizes.min_body が null のときに使う）。
DEFAULT_MIN_BODY = 10.0


def opt_size(value: Any, *, key: str = "", warn=None) -> float | None:
    """文字サイズを float に揃える。null / 空文字 / 不正値は「指定なし」。

    Dify や LLM が数値を "18" のような文字列で送ってくるので数値文字列は受ける。
    bool は弾く（Python では isinstance(True, int) が真なので、YAML の
    `title: yes` が素通りして Pt(True) = 1 EMU になってしまう）。
    """
    def reject(reason: str) -> None:
        if warn is not None:
            warn(f"sizes.{key or '?'} {reason}。この項目は指定なしとして扱います")

    if value is None:
        return None
    if isinstance(value, bool):
        reject(f"は数値で指定してください（受け取った値: {value}）")
        return None
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        try:
            value = float(text)
        except ValueError:
            reject(f"を数値として解釈できません（受け取った値: {text[:20]!r}）")
            return None
    if not isinstance(value, (int, float)):
        reject(f"は数値で指定してください（受け取った型: {type(value).__name__}）")
        return None
    size = float(value)
    if not MIN_FONT_SIZE <= size <= MAX_FONT_SIZE:
        reject(f"は {MIN_FONT_SIZE:g}〜{MAX_FONT_SIZE:g} の範囲で指定してください（{size:g}）")
        return None
    return size


def set_typeface(rPr, latin: str | None, ea: str | None) -> None:
    """<a:latin> の直後に <a:ea>/<a:cs> を差し込む（日本語が明朝に落ちるのを防ぐ）。

    None を渡した書体は要素を書き込まないので、スライドマスター（テーマ）の
    フォントがそのまま効く。
    """
    prev = None
    for tag, face in ((qn("a:latin"), latin), (qn("a:ea"), ea),
                      (qn("a:cs"), latin)):
        if face is None:
            continue
        el = rPr.find(tag)
        if el is None:
            el = rPr.makeelement(tag, {"typeface": face})
            if prev is None:
                # latin を書かない場合でも、既にある <a:solidFill> などより
                # 後ろに置く必要があるので末尾に足す。
                rPr.append(el)
            else:
                rPr.insert(list(rPr).index(prev) + 1, el)
        else:
            el.set("typeface", face)
        prev = el


# lang / noProof を持てるのはこの 3 要素だけ。
LANG_TAGS = (qn("a:rPr"), qn("a:endParaRPr"), qn("a:defRPr"))
# テンプレートに lang が 1 つも無いときに使う言語。
DEFAULT_LANGUAGE = "ja-JP"
_LANG_RE = re.compile(r"^[A-Za-z]{2,3}(-[A-Za-z0-9]{2,8})*$")


def _lang_roots(prs) -> Iterable[Any]:
    """言語判定のために走査するテンプレート側のパート。

    ノートマスターは含めない。python-pptx が英語の既定ノートマスターを
    後から注入することがあり、それを数えると判定が英語に倒れる。
    """
    yield prs._element                      # presentation.xml（defaultTextStyle）
    for master in prs.slide_masters:
        yield master._element
        # prs.slide_layouts は第 1 マスター配下しか返さないのでここで辿る。
        for layout in master.slide_layouts:
            yield layout._element
    for slide in prs.slides:                # テンプレートに元からあるスライド
        yield slide._element


def count_languages(prs) -> dict[str, int]:
    """テンプレートに書かれている lang の値と出現数。"""
    counts: dict[str, int] = {}
    for root in _lang_roots(prs):
        for el in root.iter():
            if el.tag not in LANG_TAGS:
                continue
            code = (el.get("lang") or "").strip()
            if code:
                counts[code] = counts.get(code, 0) + 1
    return counts


def detect_language(prs, fallback: str = DEFAULT_LANGUAGE) -> str:
    """テンプレートに書かれている lang の最頻値を返す。

    altLang は数えない。日本語のテンプレートは lang="ja-JP" と対で
    altLang="en-US" を持つので、数えると英語が勝ってしまう。
    """
    counts = count_languages(prs)
    if not counts:
        return fallback
    # 同数なら辞書順。同じテンプレートなら必ず同じ答えになるようにする。
    return min(counts.items(), key=lambda kv: (-kv[1], kv[0]))[0]


# ════════════════════════════════════════════════════════════════════════
# 4b. テンプレートの文字サイズ
# ════════════════════════════════════════════════════════════════════════

def _defrpr_size(el, level: int) -> float | None:
    """<...><a:lvl{level}pPr><a:defRPr sz="3200"/> を pt で返す。"""
    if el is None:
        return None
    lvl = el.find(qn(f"a:lvl{level}pPr"))
    if lvl is None:
        return None
    rpr = lvl.find(qn("a:defRPr"))
    sz = None if rpr is None else rpr.get("sz")
    return float(sz) / 100.0 if sz else None


def placeholder_size(layout, idx: int, level: int = 1) -> float | None:
    """レイアウトのプレースホルダが持つ <a:lstStyle> の文字サイズ。

    ph.text_frame.paragraphs[0].font.size では取れない。あれが見るのは段落の
    <a:pPr><a:defRPr> で、レイアウトのサイズが入っているのは <a:lstStyle> の
    ほう。スライド上のプレースホルダからも取れない（add_slide が複製するのは
    <p:ph> の型と idx だけで lstStyle は来ない）ので、必ずレイアウトを渡す。
    """
    if layout is None:
        return None
    for ph in layout.placeholders:
        if ph.placeholder_format.idx != idx:
            continue
        body = ph._element.find(qn("p:txBody"))
        return _defrpr_size(None if body is None else body.find(qn("a:lstStyle")), level)
    return None


def master_style_size(master, style: str, level: int = 1) -> float | None:
    """スライドマスターの <p:txStyles><p:{title|body|other}Style> の文字サイズ。"""
    if master is None:
        return None
    styles = master._element.find(qn("p:txStyles"))
    return _defrpr_size(None if styles is None else styles.find(qn(f"p:{style}")), level)


def default_text_size(prs, level: int = 1) -> float | None:
    """presentation.xml の <p:defaultTextStyle>。素の textbox の継承元。"""
    return _defrpr_size(prs._element.find(qn("p:defaultTextStyle")), level)


# テンプレートに対応する書式スロットが無いキーは、そのコンテキストの本文
# サイズ（lvl1）を基準に、既定値が持っていた比率で算出して書き込む。既定値は
# 暗黙に「基準 18pt」で調律されていたので、基準が 18 のコンテキスト（素の
# テキストボックス）では今までとまったく同じ数字になる。
_DERIVED_RATIO = {
    "h3": 20 / 18,
    "h4": 17 / 18,
    "code": 12 / 18,
    "caption": 11 / 18,
    "page_number": 10 / 18,
}
# テンプレートから 1 つも読めなかったときの最終フォールバック。
FALLBACK_FONT_SIZE = 18.0


class SizeBook:
    """文字サイズの「書き込む値」と「見積りに使う値」を分けて持つ。

    write()  … run に書き込む値。None なら sz を書かない（テンプレート継承）。
    metric() … 高さ見積りに使う値。必ず float を返す。

    同じ null でも、描かれる場所によって PowerPoint が適用する継承元が違うので
    コンテキスト（ctx）を必ず渡す。
      "placeholder" … 本文プレースホルダ（レイアウトの ph → マスター bodyStyle）
      "title"       … タイトル／サブタイトル（レイアウトの ph → titleStyle）
      "shape"       … 自前の textbox・図形・表（defaultTextStyle → otherStyle）
    """

    def __init__(self, cfg: dict, prs, layout_of, warn) -> None:
        self.cfg = cfg
        self.prs = prs
        self.layout_of = layout_of        # Renderer.layout（レイアウトキー → layout）
        self.warn = warn
        raw = cfg["sizes"]
        self.user = {k: opt_size(v, key=k, warn=warn)
                     for k, v in raw.items() if k != "body"}
        self.user["body"] = self._normalize_body(raw.get("body"))
        self._cache: dict[tuple, float] = {}

    # ---------- 公開 API ----------

    def write(self, key: str, *, level: int = 0, ctx: str = "shape",
              layout_key: str | None = None, scale: float = 1.0) -> float | None:
        """run に書き込むサイズ。None なら sz を書かない。"""
        size = self._user_value(key, level)
        if size is None and key in _DERIVED_RATIO:
            size = self._derive(key, ctx, layout_key)
        if size is None:
            if scale >= 1.0:
                return None               # 本来の null。テンプレートに任せる
            # 縮小は sz を書かないと実現できない（OOXML に相対指定が無い）。
            size = self._template(key, level, ctx, layout_key)
        return self._clamp(size * scale)

    def metric(self, key: str, *, level: int = 0, ctx: str = "shape",
               layout_key: str | None = None, scale: float = 1.0) -> float:
        """高さ見積りに使うサイズ。書き込む値があるなら必ずそれと一致する。"""
        written = self.write(key, level=level, ctx=ctx,
                             layout_key=layout_key, scale=scale)
        if written is not None:
            return written
        return self._clamp(self._template(key, level, ctx, layout_key) * scale)

    def floor(self) -> float:
        """自動縮小の下限（sizes.min_body。null なら組み込みの既定）。"""
        return self.user.get("min_body") or DEFAULT_MIN_BODY

    # ---------- 設定値 ----------

    def _normalize_body(self, value) -> list[float | None] | None:
        """body は「リスト / スカラー / null」のどれでも受ける。"""
        if value is None:
            return None
        if not isinstance(value, list):
            value = [value]
        sizes = [opt_size(v, key="body", warn=self.warn) for v in value]
        return sizes or None

    def _user_value(self, key: str, level: int) -> float | None:
        if key != "body":
            return self.user.get(key)
        sizes = self.user.get("body")
        if not sizes:
            return None
        return sizes[min(level, len(sizes) - 1)]

    # ---------- テンプレート解決 ----------

    def _derive(self, key: str, ctx: str, layout_key: str | None) -> float:
        base = self._template("body", 0, ctx, layout_key)
        size = base * _DERIVED_RATIO[key]
        if key in ("h3", "h4"):
            # 小見出しがタイトルを追い越さないように抑える。
            title = self._template("title", 0, "title", layout_key)
            size = min(size, title * 0.9)
            if key == "h4":
                size = min(size, self._derive("h3", ctx, layout_key))
        return self._clamp(size)

    def _template(self, key: str, level: int, ctx: str,
                  layout_key: str | None) -> float:
        """テンプレートに書かれている実効サイズ。必ず float を返す。"""
        if key in _DERIVED_RATIO:
            return self._derive(key, ctx, layout_key)
        cache_key = (key, level, ctx, layout_key)
        if cache_key not in self._cache:
            self._cache[cache_key] = self._resolve(key, level, ctx, layout_key)
        return self._cache[cache_key]

    def _resolve(self, key: str, level: int, ctx: str,
                 layout_key: str | None) -> float:
        phs = self.cfg["placeholders"]
        layout = self.layout_of(layout_key) if layout_key else None
        master = getattr(layout, "slide_master", None)
        if master is None and self.prs.slide_masters:
            master = self.prs.slide_masters[0]
        # 本文プレースホルダは lvl5 までしか使わない（fill_text_frame が
        # para.level を 4 でクランプしている）ので、解決レベルも揃える。
        lvl = min(level, 4) + 1

        if key == "cover_subtitle":
            chain = (placeholder_size(layout, phs["subtitle"], 1),
                     master_style_size(master, "bodyStyle", 1),
                     default_text_size(self.prs, 1))
        elif ctx == "title":
            chain = (placeholder_size(layout, phs["title"], 1),
                     master_style_size(master, "titleStyle", 1),
                     default_text_size(self.prs, 1))
        elif ctx == "placeholder":
            chain = (placeholder_size(layout, phs["body"], lvl),
                     master_style_size(master, "bodyStyle", lvl),
                     default_text_size(self.prs, lvl))
        else:                              # shape: 自前の textbox・図形・表
            chain = (default_text_size(self.prs, lvl),
                     master_style_size(master, "otherStyle", lvl),
                     default_text_size(self.prs, 1))
        for size in chain:
            if size:
                return self._clamp(size)
        return FALLBACK_FONT_SIZE

    @staticmethod
    def _clamp(size: float) -> float:
        return max(MIN_FONT_SIZE, min(MAX_FONT_SIZE, float(size)))


class Painter:
    """設定に基づいて run / 段落を装飾するヘルパ。"""

    def __init__(self, cfg: dict):
        self.cfg = cfg
        # 空欄（null / 空文字）は「指定しない」= マスターのフォントに任せる。
        self.f = {k: opt_str(v) for k, v in cfg["fonts"].items()}
        # 色も同じ。null の項目は塗らないので、テンプレートの配色が残る。
        self.c = {k: opt_str(v) for k, v in cfg["colors"].items()}

    def style_run(self, run, *, size: float | None, bold=False, italic=False,
                  color: str | None = None, mono=False, strike=False,
                  link: str | None = None) -> None:
        font = run.font
        # None は「サイズを指定しない」。sz 属性が消えてテンプレートを継承する。
        font.size = None if size is None else Pt(size)
        font.bold = bold
        font.italic = italic
        if color:
            font.color.rgb = rgb(color)
        latin = self.f["code"] if mono else self.f["latin"]
        ea = self.f["code_eastasian"] if mono else self.f["eastasian"]
        set_typeface(run._r.get_or_add_rPr(), latin, ea)
        if strike:
            run._r.get_or_add_rPr().set("strike", "sngStrike")
        if link:
            try:
                run.hyperlink.address = link
            except Exception:
                pass

    def fill_runs(self, para, runs: Sequence[Run], *, size: float | None,
                  color: str | None = None, base_bold=False,
                  base_italic=False) -> None:
        """Run 列を 1 段落に流し込む（改行 Run はソフト改行として扱う）。

        color=None は「文字色を指定しない」。呼び出し側は colors の値を
        そのまま渡すので、null にした色はここで書き込まれず継承色になる。
        """
        if not runs:
            para.add_run().text = ""
            return
        for r in runs:
            if r.text == "\n":
                br = para._p.makeelement(qn("a:br"), {})
                para._p.append(br)
                continue
            run = para.add_run()
            run.text = r.text
            self.style_run(
                run,
                # サイズ未指定のときはインラインコードの 6% 縮小も効かない
                # （sz を書かずに相対縮小する手段が OOXML に無い）。
                size=None if size is None else size * (0.94 if r.code else 1.0),
                bold=base_bold or r.bold,
                italic=base_italic or r.italic,
                color=self.c["accent"] if r.link else (self.c["code_text"] if r.code else color),
                mono=r.code,
                strike=r.strike,
                link=r.link,
            )

    def indent(self, para, level: int, hanging: bool) -> None:
        """marL / indent を明示指定する（テンプレート依存の字下げブレを防ぐ）。"""
        sp = self.cfg["spacing"]
        step, hang = sp.get("list_indent"), sp.get("list_hanging")
        if step is None:
            return
        pPr = para._p.get_or_add_pPr()
        if hanging:
            pPr.set("marL", str(int(step * (level + 1) * EMU_PER_IN)))
            pPr.set("indent", str(int(-(hang or step) * EMU_PER_IN)))
        else:
            pPr.set("marL", str(int(step * level * EMU_PER_IN)))
            pPr.set("indent", "0")

    def no_bullet(self, para) -> None:
        pPr = para._p.get_or_add_pPr()
        for tag in ("a:buChar", "a:buAutoNum", "a:buNone"):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))

    def bullet_char(self, para, char: str = "•") -> None:
        pPr = para._p.get_or_add_pPr()
        for tag in ("a:buChar", "a:buAutoNum", "a:buNone"):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)
        font = pPr.makeelement(qn("a:buFont"), {"typeface": "Arial"})
        pPr.append(font)
        pPr.append(pPr.makeelement(qn("a:buChar"), {"char": char}))

    def auto_number(self, para, fmt: str = "arabicPeriod") -> None:
        pPr = para._p.get_or_add_pPr()
        for tag in ("a:buChar", "a:buAutoNum", "a:buNone"):
            el = pPr.find(qn(tag))
            if el is not None:
                pPr.remove(el)
        pPr.append(pPr.makeelement(qn("a:buAutoNum"), {"type": fmt}))


# ════════════════════════════════════════════════════════════════════════
# 5. レンダラ
# ════════════════════════════════════════════════════════════════════════

BULLET_CHARS = ["•", "–", "‣", "·", "·"]


class Renderer:
    def __init__(self, cfg: dict, base_dir: Path, config_dir: Path | None = None):
        self.cfg = cfg
        self.warnings: list[str] = []
        self.base_dir = base_dir          # 画像の相対パス基準（Markdown のあるフォルダ）
        self.paint = Painter(cfg)
        self.c = self.paint.c             # null を None に均した配色
        self.prs = self._open_template(self._find_template(cfg["template"], config_dir, base_dir))
        # 言語はスライドを 1 枚も足す前に見る。attach_notes() が notes_slide に
        # 触ると python-pptx が英語の既定ノートマスターを注入してしまうため。
        self.template_lang = detect_language(self.prs)
        self.applied_language: str | None = None
        self.layouts = {l.name: l for l in self.prs.slide_layouts}
        self.slide_w_pt = self.prs.slide_width / EMU_PER_PT
        self.slide_h_pt = self.prs.slide_height / EMU_PER_PT
        self.sizes = SizeBook(cfg, self.prs, self.layout, self.warnings.append)

    # ---------- テンプレート ----------

    @staticmethod
    def _find_template(name: str, config_dir: Path | None, base_dir: Path) -> Path:
        p = Path(name)
        if p.is_absolute():
            return p
        for d in [d for d in (config_dir, base_dir, Path.cwd()) if d is not None]:
            if (d / p).exists():
                return d / p
        return p

    @staticmethod
    def _open_template(path: Path) -> Presentation:
        if not path.exists():
            raise SystemExit(f"テンプレートが見つかりません: {path}")
        if path.suffix.lower() == ".potx":
            tmp = Path(tempfile.mkdtemp()) / (path.stem + ".pptx")
            shutil.copy(path, tmp)
            return Presentation(tmp)
        return Presentation(path)

    def layout(self, key: str):
        name = self.cfg["layouts"].get(key)
        if name in self.layouts:
            return self.layouts[name]
        self.warnings.append(f"レイアウト '{name}'（{key}）が見つからないため既定を使用します")
        return self.prs.slide_layouts[0]

    def add_slide(self, key: str):
        return self.prs.slides.add_slide(self.layout(key))

    # ---------- プレースホルダ ----------

    @staticmethod
    def _ph(slide, idx: int):
        for ph in slide.placeholders:
            if ph.placeholder_format.idx == idx:
                return ph
        return None

    def title_ph(self, slide):
        ph = self._ph(slide, self.cfg["placeholders"]["title"])
        if ph is not None:
            return ph
        for p in slide.placeholders:
            if p.placeholder_format.type in (13, 15, 16):
                continue
            if "title" in p.name.lower() or "タイトル" in p.name:
                return p
        return None

    @staticmethod
    def drop(shape) -> None:
        shape._element.getparent().remove(shape._element)

    def clear_empty_placeholders(self, slide, keep: set[int]) -> None:
        for ph in list(slide.placeholders):
            idx = ph.placeholder_format.idx
            if idx in keep:
                continue
            if ph.placeholder_format.type in (13, 15, 16):  # 番号/フッタ/日付
                continue
            if ph.has_text_frame and not ph.text_frame.text.strip():
                self.drop(ph)

    # ---------- 領域 ----------

    def _layout_placeholder(self, layout_key: str, idx: int):
        """レイアウト（スライドではない）のプレースホルダを idx で引く。"""
        for ph in self.layout(layout_key).placeholders:
            if ph.placeholder_format.idx == idx:
                return ph
        return None

    def has_body_placeholder(self) -> bool:
        """本文プレースホルダを持つレイアウトか。描画経路と ctx を決める。"""
        return self._layout_placeholder(
            "content", self.cfg["placeholders"]["body"]) is not None

    def content_rect_pt(self) -> tuple[float, float, float, float]:
        ba = self.cfg["body_area"]
        if all(ba.get(k) is not None for k in ("left", "top", "width", "height")):
            return (ba["left"] * 72, ba["top"] * 72, ba["width"] * 72, ba["height"] * 72)
        ph = self._layout_placeholder("content", self.cfg["placeholders"]["body"])
        if ph is not None and ph.left is not None:
            return (ph.left / EMU_PER_PT, ph.top / EMU_PER_PT,
                    ph.width / EMU_PER_PT, ph.height / EMU_PER_PT)
        # フォールバック: 上下左右にマージン
        m = 0.7 * 72
        top = 1.6 * 72
        return (m, top, self.slide_w_pt - 2 * m, self.slide_h_pt - top - 0.7 * 72)

    # ---------- 高さ見積り ----------

    def measure(self, blk: Block, width_pt: float, scale: float = 1.0,
                ctx: str = "shape") -> float:
        sp = self.cfg["spacing"]
        lr = sp["line_ratio"]
        lay = "content" if ctx == "placeholder" else "table"

        def size_of(key, level=0):
            return self.sizes.metric(key, level=level, ctx=ctx,
                                     layout_key=lay, scale=scale)

        if isinstance(blk, Heading):
            size = size_of("h3" if blk.level <= 3 else "h4")
            return wrapped_lines(runs_text(blk.runs), size, width_pt) * size * lr + 6
        if isinstance(blk, Para):
            size = size_of("paragraph")
            return wrapped_lines(runs_text(blk.runs), size, width_pt) * size * lr + sp["para_space_after"]
        if isinstance(blk, ListBlock):
            total = 0.0
            for it in blk.items:
                size = size_of("body", it.level)
                indent = 18 * (it.level + 1)
                total += wrapped_lines(runs_text(it.runs), size, width_pt - indent) * size * lr
                total += sp["para_space_after"] * 0.6
            return total
        if isinstance(blk, Table):
            return self.table_height(blk, width_pt, scale)
        if isinstance(blk, Code):
            size = self.sizes.metric("code", ctx="shape", scale=scale)
            lines = sum(max(1, math.ceil(char_units(l) * 0.9 / max(width_pt / size, 1)))
                        for l in blk.text.split("\n")) or 1
            return lines * size * 1.28 + 16
        if isinstance(blk, Quote):
            inner = width_pt - self.cfg["quote"]["indent"] * 72
            return sum(self.measure(b, inner, scale, ctx) for b in blk.blocks) + 8
        if isinstance(blk, Image):
            return self.image_size_pt(blk, width_pt)[1] + 6
        if isinstance(blk, Rule):
            return 12
        return 0.0

    def table_height(self, tb: Table, width_pt: float, scale: float = 1.0) -> float:
        # 表は常に自前配置（手動レイアウト）なので ctx は shape 固定。
        size = self.sizes.metric("table", ctx="shape", scale=scale)
        cols = self.table_widths(tb, width_pt)
        h = 0.0
        rows = ([tb.header] if tb.header else []) + tb.rows
        for row in rows:
            lines = 1
            for i, cell in enumerate(row):
                w = cols[i] if i < len(cols) else width_pt / max(len(row), 1)
                pad = self.cfg["table"]["cell_margin"] * 72 * 2
                lines = max(lines, wrapped_lines(runs_text(cell), size, w - pad))
            h += max(lines * size * 1.3 + 8, self.cfg["table"]["min_row_height"] * 72)
        return h + 4

    def table_widths(self, tb: Table, width_pt: float) -> list[float]:
        rows = ([tb.header] if tb.header else []) + tb.rows
        ncol = max((len(r) for r in rows), default=1)
        weights = [1.0] * ncol
        for r in rows:
            for i, cell in enumerate(r[:ncol]):
                weights[i] = max(weights[i], min(char_units(runs_text(cell)), 28.0))
        total = sum(weights) or 1
        raw = [max(width_pt * w / total, 46.0) for w in weights]
        scale = width_pt / sum(raw)
        return [w * scale for w in raw]

    def image_size_pt(self, img: Image, width_pt: float) -> tuple[float, float]:
        path = self.resolve_image(img.src)
        max_h = self.cfg["image"]["max_height"] * 72
        if path is None:
            return (width_pt, 40.0)
        try:
            from PIL import Image as PILImage
            with PILImage.open(path) as im:
                w, h = im.size
            ratio = h / w
        except Exception:
            ratio = 0.62
        w_pt = width_pt
        h_pt = w_pt * ratio
        if h_pt > max_h:
            h_pt = max_h
            w_pt = h_pt / ratio
        return (w_pt, h_pt)

    def resolve_image(self, src: str) -> Path | None:
        if re.match(r"^[a-z]+://", src, re.I):
            return None
        p = Path(src)
        if not p.is_absolute():
            p = self.base_dir / p
        return p if p.exists() else None

    # ---------- 描画 ----------

    def render(self, specs: list[SlideSpec]) -> None:
        for spec in specs:
            if spec.kind == "cover":
                self.render_cover(spec)
            elif spec.kind == "section":
                self.render_section(spec)
            else:
                self.render_content(spec)
        if self.cfg["options"]["page_number"]:
            self.add_page_numbers()
        # ページ番号の run も対象にしたいので必ず最後。
        self.apply_language()

    def set_title(self, slide, runs: Sequence[Run], key: str, layout_key: str,
                  color: str | None = None) -> None:
        """タイトルを流し込む。

        key はサイズ設定のキー、layout_key はそのスライドが使うレイアウト。
        レイアウトごとにタイトルの文字サイズが違うことがあるので、どちらも要る。
        """
        ph = self.title_ph(slide)
        if ph is None or not runs:
            if ph is not None and not runs:
                self.drop(ph)
            return
        tf = ph.text_frame
        tf.word_wrap = True
        para = tf.paragraphs[0]
        text = runs_text(runs)
        width_pt = (ph.width or Inches(10)) / EMU_PER_PT
        height_pt = (ph.height or Inches(1.2)) / EMU_PER_PT
        # タイトルが長い場合は自動で縮小する。縮める必要が無ければ、設定どおり
        # （既定なら null ＝ サイズを書かず、テンプレートのタイトル書式に任せる）。
        base = self.sizes.metric(key, ctx="title", layout_key=layout_key)
        size = base
        floor = max(self.sizes.floor(), 14.0)
        while size > floor and wrapped_lines(text, size, width_pt) * size * 1.2 > height_pt:
            size -= 2
        written = (size if size < base
                   else self.sizes.write(key, ctx="title", layout_key=layout_key))
        self.paint.fill_runs(para, runs, size=written,
                             color=color or self.c["heading"], base_bold=True)
        self.paint.no_bullet(para)

    def render_cover(self, spec: SlideSpec) -> None:
        slide = self.add_slide("title")
        self.set_title(slide, spec.title, "cover_title", "title")
        sub = self._ph(slide, self.cfg["placeholders"]["subtitle"])
        if sub is not None:
            if spec.subtitle:
                tf = sub.text_frame
                tf.word_wrap = True
                para = tf.paragraphs[0]
                self.paint.fill_runs(para, spec.subtitle,
                                     size=self.sizes.write("cover_subtitle",
                                                           ctx="title",
                                                           layout_key="title"),
                                     color=self.c["muted"])
                self.paint.no_bullet(para)
            else:
                self.drop(sub)
        self.clear_empty_placeholders(slide, {self.cfg["placeholders"]["title"]})
        self.attach_notes(slide, spec)

    def render_section(self, spec: SlideSpec) -> None:
        slide = self.add_slide("section")
        self.set_title(slide, spec.title, "section_title", "section")
        self.clear_empty_placeholders(slide, {self.cfg["placeholders"]["title"]})
        self.attach_notes(slide, spec)

    # --- 本文スライド ---

    @staticmethod
    def is_text_only(blocks: Sequence[Block]) -> bool:
        return all(isinstance(b, (Heading, Para, ListBlock)) for b in blocks) and bool(blocks)

    def render_content(self, spec: SlideSpec) -> None:
        opts = self.cfg["options"]
        if not spec.blocks:
            slide = self.add_slide("content")
            self.set_title(slide, spec.title, "title", "content")
            self.clear_empty_placeholders(slide, {self.cfg["placeholders"]["title"]})
            self.attach_notes(slide, spec)
            return

        # 経路はチャンクごとではなく spec ごとに決める。チャンク単位で決めると、
        # 表を含む spec の中のテキストだけのチャンクが本文プレースホルダへ流れる
        # のに、見積りは手動レイアウト前提のままになり、文字サイズの継承元が
        # 食い違ってあふれる。
        placeholder = self.is_text_only(spec.blocks) and self.has_body_placeholder()
        ctx = "placeholder" if placeholder else "shape"
        chunks = self.paginate(spec.blocks, ctx)
        for n, (blocks, scale) in enumerate(chunks):
            title = list(spec.title)
            if n > 0 and title:
                title = title + [Run(opts["continuation_suffix"])]
            slide = (self.render_text_slide(title, blocks, scale) if placeholder
                     else self.render_rich_slide(title, blocks, scale))
            if n == 0:
                self.attach_notes(slide, spec)

    def paginate(self, blocks: list[Block],
                 ctx: str = "shape") -> list[tuple[list[Block], float]]:
        """収まるように分割し、(ブロック列, フォント倍率) を返す。"""
        opts = self.cfg["options"]
        left, top, width, height = self.content_rect_pt()
        gap = self.cfg["spacing"]["block_gap"]

        # まず全体が入るか、段階的に縮小しながら確認
        for step in range(opts["shrink_steps"] + 1):
            scale = 1.0 - 0.08 * step
            total = sum(self.measure(b, width, scale, ctx) for b in blocks) + gap * (len(blocks) - 1)
            if total <= height:
                return [(blocks, scale)]
        if not opts["auto_split"]:
            return [(blocks, 1.0 - 0.08 * opts["shrink_steps"])]

        scale = 1.0 - 0.08 * opts["shrink_steps"]
        chunks: list[tuple[list[Block], float]] = []
        cur: list[Block] = []
        used = 0.0
        for blk in blocks:
            pieces = self.split_block(blk, width, height, scale, ctx)
            for piece in pieces:
                h = self.measure(piece, width, scale, ctx)
                if cur and used + gap + h > height:
                    chunks.append((cur, scale))
                    cur, used = [], 0.0
                cur.append(piece)
                used += h + gap
        if cur:
            chunks.append((cur, scale))
        return chunks or [(blocks, scale)]

    def split_block(self, blk: Block, width: float, height: float, scale: float,
                    ctx: str = "shape") -> list[Block]:
        """単体でスライドに収まらないブロックを分割する（表とリスト）。"""
        if self.measure(blk, width, scale, ctx) <= height:
            return [blk]
        if isinstance(blk, Table):
            out: list[Table] = []
            cur = Table(header=blk.header, rows=[], aligns=blk.aligns)
            for row in blk.rows:
                trial = Table(header=blk.header, rows=cur.rows + [row], aligns=blk.aligns)
                if cur.rows and self.table_height(trial, width, scale) > height:
                    out.append(cur)
                    cur = Table(
                        header=blk.header if self.cfg["options"]["table_header_repeat"] else [],
                        rows=[row], aligns=blk.aligns)
                else:
                    cur = trial
            if cur.rows:
                out.append(cur)
            return out or [blk]
        if isinstance(blk, ListBlock):
            out: list[ListBlock] = []
            cur = ListBlock()
            for it in blk.items:
                trial = ListBlock(cur.items + [it])
                if cur.items and self.measure(trial, width, scale, ctx) > height:
                    out.append(cur)
                    cur = ListBlock([it])
                else:
                    cur = trial
            if cur.items:
                out.append(cur)
            return out or [blk]
        return [blk]

    def render_text_slide(self, title: list[Run], blocks: list[Block], scale: float):
        """テキストだけのスライド → 本文プレースホルダに流す（テンプレート書式を継承）。"""
        slide = self.add_slide("content")
        self.set_title(slide, title, "title", "content")
        body = self._ph(slide, self.cfg["placeholders"]["body"])
        if body is None:
            self.clear_empty_placeholders(slide, {self.cfg["placeholders"]["title"]})
            left, top, width, height = self.content_rect_pt()
            self.flow_blocks(slide, blocks, left, top, width, height, scale)
            return slide
        tf = body.text_frame
        tf.word_wrap = True
        self.fill_text_frame(tf, blocks, scale, ctx="placeholder")
        self.clear_empty_placeholders(
            slide, {self.cfg["placeholders"]["title"], self.cfg["placeholders"]["body"]})
        return slide

    def render_rich_slide(self, title: list[Run], blocks: list[Block], scale: float):
        """表・コード・画像を含むスライド → 手動レイアウト。"""
        slide = self.add_slide("table")
        self.set_title(slide, title, "title", "table")
        self.clear_empty_placeholders(slide, {self.cfg["placeholders"]["title"]})
        left, top, width, height = self.content_rect_pt()
        self.flow_blocks(slide, blocks, left, top, width, height, scale)
        return slide

    # --- テキストフレームへの流し込み ---

    def fill_text_frame(self, tf, blocks: Sequence[Block], scale: float,
                        first_para=None, ctx: str = "shape") -> None:
        p = self.cfg
        lay = "content" if ctx == "placeholder" else "table"
        started = False

        def size_of(key, level=0):
            return self.sizes.write(key, level=level, ctx=ctx,
                                    layout_key=lay, scale=scale)

        def new_para():
            nonlocal started
            if not started:
                started = True
                return tf.paragraphs[0]
            return tf.add_paragraph()

        for blk in blocks:
            if isinstance(blk, Heading):
                para = new_para()
                size = size_of("h3" if blk.level <= 3 else "h4")
                self.paint.fill_runs(para, blk.runs, size=size,
                                     color=self.c["accent"], base_bold=True)
                self.paint.no_bullet(para)
                self.paint.indent(para, 0, hanging=False)
                para.space_before = Pt(8 if started else 0)
                para.space_after = Pt(3)
            elif isinstance(blk, Para):
                para = new_para()
                self.paint.fill_runs(para, blk.runs, size=size_of("paragraph"),
                                     color=self.c["text"])
                self.paint.no_bullet(para)
                self.paint.indent(para, 0, hanging=False)
                para.space_after = Pt(p["spacing"]["para_space_after"])
            elif isinstance(blk, ListBlock):
                for it in blk.items:
                    para = new_para()
                    para.level = min(it.level, 4)
                    self.paint.fill_runs(para, it.runs,
                                         size=size_of("body", it.level),
                                         color=self.c["text"])
                    if it.ordered:
                        self.paint.auto_number(para)
                    else:
                        self.paint.bullet_char(para, BULLET_CHARS[min(it.level, 4)])
                    self.paint.indent(para, it.level, hanging=True)
                    para.space_after = Pt(p["spacing"]["para_space_after"] * 0.6)
            elif isinstance(blk, Quote):
                for sub in blk.blocks:
                    para = new_para()
                    runs = sub.runs if isinstance(sub, (Para, Heading)) else []
                    self.paint.fill_runs(para, runs, size=size_of("quote"),
                                         color=self.c["quote_text"],
                                         base_italic=p["quote"]["italic"])
                    self.paint.no_bullet(para)
                    self.paint.indent(para, 0, hanging=False)
                    para.space_after = Pt(p["spacing"]["para_space_after"])

    # --- 手動フロー ---

    def flow_blocks(self, slide, blocks: Sequence[Block], left: float, top: float,
                    width: float, height: float, scale: float) -> None:
        y = top
        gap = self.cfg["spacing"]["block_gap"]
        for blk in blocks:
            # 手動フローで描く図形は常にテンプレートの既定テキスト書式を継承する。
            h = self.measure(blk, width, scale, "shape")
            self.draw_block(slide, blk, left, y, width, h, scale)
            y += h + gap

    def draw_block(self, slide, blk: Block, x: float, y: float,
                   w: float, h: float, scale: float) -> None:
        if isinstance(blk, Table):
            self.draw_table(slide, blk, x, y, w, h, scale)
        elif isinstance(blk, Code):
            self.draw_code(slide, blk, x, y, w, h, scale)
        elif isinstance(blk, Image):
            self.draw_image(slide, blk, x, y, w, scale)
        elif isinstance(blk, Quote):
            self.draw_quote(slide, blk, x, y, w, h, scale)
        elif isinstance(blk, Rule):
            self.draw_rule(slide, x, y + h / 2, w)
        else:
            box = self.textbox(slide, x, y, w, h)
            self.fill_text_frame(box.text_frame, [blk], scale)

    def textbox(self, slide, x: float, y: float, w: float, h: float):
        box = slide.shapes.add_textbox(Pt(x), Pt(y), Pt(w), Pt(max(h, 12)))
        tf = box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        tf.vertical_anchor = MSO_ANCHOR.TOP
        return box

    def draw_rule(self, slide, x: float, y: float, w: float) -> None:
        from pptx.enum.shapes import MSO_CONNECTOR
        line = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Pt(x), Pt(y), Pt(x + w), Pt(y))
        if self.c["table_border"]:
            line.line.color.rgb = rgb(self.c["table_border"])
        line.line.width = Pt(1)
        line.shadow.inherit = False

    def draw_quote(self, slide, q: Quote, x: float, y: float, w: float,
                   h: float, scale: float) -> None:
        indent = self.cfg["quote"]["indent"] * 72
        if self.cfg["quote"]["bar"]:
            from pptx.enum.shapes import MSO_SHAPE
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(3), Pt(h))
            if self.c["quote_bar"]:       # null ならテーマの図形色のまま
                bar.fill.solid()
                bar.fill.fore_color.rgb = rgb(self.c["quote_bar"])
            bar.line.fill.background()
            bar.shadow.inherit = False
        box = self.textbox(slide, x + indent, y, w - indent, h)
        self.fill_text_frame(box.text_frame, q.blocks, scale)
        for para in box.text_frame.paragraphs:
            for run in para.runs:
                run.font.italic = self.cfg["quote"]["italic"]
                if self.c["quote_text"]:
                    run.font.color.rgb = rgb(self.c["quote_text"])

    def draw_code(self, slide, code: Code, x: float, y: float, w: float,
                  h: float, scale: float) -> None:
        from pptx.enum.shapes import MSO_SHAPE
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
        shape.adjustments[0] = 0.04
        # null の色は書き込まない＝テーマの図形書式（塗り・線）がそのまま残る。
        if self.c["code_bg"]:
            shape.fill.solid()
            shape.fill.fore_color.rgb = rgb(self.c["code_bg"])
        if self.c["table_border"]:
            shape.line.color.rgb = rgb(self.c["table_border"])
        shape.line.width = Pt(0.75)
        shape.shadow.inherit = False
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(9)
        tf.margin_top = tf.margin_bottom = Pt(7)
        tf.vertical_anchor = MSO_ANCHOR.TOP
        size = self.sizes.write("code", ctx="shape", scale=scale)
        lines = code.text.split("\n")
        for i, line in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = PP_ALIGN.LEFT
            self.paint.no_bullet(para)
            run = para.add_run()
            run.text = line if line else " "
            self.paint.style_run(run, size=size, color=self.c["code_text"], mono=True)
            para.line_spacing = 1.0
            para.space_after = Pt(0)

    def draw_image(self, slide, img: Image, x: float, y: float, w: float, scale: float) -> None:
        path = self.resolve_image(img.src)
        if path is None:
            box = self.textbox(slide, x, y, w, 20)
            para = box.text_frame.paragraphs[0]
            self.paint.fill_runs(para, [Run(f"[画像が見つかりません: {img.src}]")],
                                 size=self.sizes.write("caption", ctx="shape"),
                                 color=self.c["muted"])
            self.paint.no_bullet(para)
            self.warnings.append(f"画像が見つかりません: {img.src}")
            return
        w_pt, h_pt = self.image_size_pt(img, w)
        off = (w - w_pt) / 2 if self.cfg["image"]["align"] == "center" else 0
        slide.shapes.add_picture(str(path), Pt(x + off), Pt(y), Pt(w_pt), Pt(h_pt))

    # --- 表 ---

    def draw_table(self, slide, tb: Table, x: float, y: float, w: float,
                   h: float, scale: float) -> None:
        rows = ([tb.header] if tb.header else []) + tb.rows
        if not rows:
            return
        ncol = max(len(r) for r in rows)
        widths = self.table_widths(tb, w)
        gfx = slide.shapes.add_table(len(rows), ncol, Pt(x), Pt(y), Pt(w), Pt(h))
        table = gfx.table
        cfgt = self.cfg["table"]
        has_header = bool(tb.header)

        table.first_row = has_header and cfgt["first_row_header"]
        table.horz_banding = bool(cfgt["banding"])
        if cfgt.get("style_id"):
            self.set_table_style(table, cfgt["style_id"])

        for i, wpt in enumerate(widths[:ncol]):
            table.columns[i].width = Emu(int(wpt * EMU_PER_PT))

        size = self.sizes.write("table", ctx="shape", scale=scale)
        head_size = self.sizes.write("table_header", ctx="shape", scale=scale)
        margin = Pt(cfgt["cell_margin"] * 72)

        for ri, row in enumerate(rows):
            is_head = has_header and ri == 0
            table.rows[ri].height = Emu(int(cfgt["min_row_height"] * EMU_PER_IN))
            for ci in range(ncol):
                cell = table.cell(ri, ci)
                cell.margin_left = cell.margin_right = margin
                cell.margin_top = cell.margin_bottom = Pt(3)
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                if cfgt["explicit_format"]:
                    self.fill_cell(cell, ri, is_head)
                tf = cell.text_frame
                tf.word_wrap = True
                para = tf.paragraphs[0]
                self.paint.no_bullet(para)
                runs = row[ci] if ci < len(row) else []
                self.paint.fill_runs(
                    para, runs,
                    size=head_size if is_head else size,
                    color=self.c["table_header_text"] if is_head
                    else self.c["table_text"],
                    base_bold=is_head,
                )
                align = tb.aligns[ci] if ci < len(tb.aligns) else "left"
                para.alignment = {"center": PP_ALIGN.CENTER,
                                  "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)

    def fill_cell(self, cell, ri: int, is_head: bool) -> None:
        """明示的にセルを塗る。色が null のセルは塗らず表スタイルの書式を残す。"""
        if is_head:
            color = self.c["table_header_bg"]
        elif self.cfg["table"]["banding"] and ri % 2 == 0:
            color = self.c["table_band_bg"]
        else:
            color = self.c["table_body_bg"]
        if not color:
            return
        cell.fill.solid()
        cell.fill.fore_color.rgb = rgb(color)

    @staticmethod
    def set_table_style(table, style_id: str) -> None:
        tbl_pr = table._tbl.find(qn("a:tblPr"))
        if tbl_pr is None:
            return
        el = tbl_pr.find(qn("a:tableStyleId"))
        if el is None:
            el = tbl_pr.makeelement(qn("a:tableStyleId"), {})
            tbl_pr.append(el)
        el.text = style_id

    # --- 付帯 ---

    def attach_notes(self, slide, spec: SlideSpec) -> None:
        if not spec.notes:
            return
        slide.notes_slide.notes_text_frame.text = "\n".join(spec.notes)

    def add_page_numbers(self) -> None:
        opts = self.cfg["options"]
        size = self.sizes.write("page_number", ctx="shape")
        for i, slide in enumerate(self.prs.slides, start=1):
            if opts["page_number_skip_first"] and i == 1:
                continue
            box = slide.shapes.add_textbox(
                Emu(self.prs.slide_width - Inches(1.3)),
                Emu(self.prs.slide_height - Inches(0.55)),
                Inches(0.9), Inches(0.32))
            tf = box.text_frame
            tf.word_wrap = False
            para = tf.paragraphs[0]
            para.alignment = PP_ALIGN.RIGHT
            self.paint.no_bullet(para)
            run = para.add_run()
            run.text = str(i)
            self.paint.style_run(run, size=size, color=self.c["muted"])

    # --- 校正言語 ---

    def resolve_language(self) -> str | None:
        """options.language を実際に書き込む値にする。None なら書かない。"""
        raw = self.cfg["options"].get("language", "auto")
        if isinstance(raw, str) and raw.strip().lower() == "auto":
            return self.template_lang
        if raw is not None and not isinstance(raw, str):
            self.warnings.append(
                f"options.language は言語タグの文字列で指定してください"
                f"（受け取った値: {type(raw).__name__}）。"
                f"テンプレートから判定した '{self.template_lang}' を使います"
            )
            return self.template_lang
        code = opt_str(raw)                  # null / 空文字は「書き込まない」
        if code is None:
            return None
        if not _LANG_RE.match(code):
            self.warnings.append(
                f"options.language '{code}' は言語タグとして解釈できません。"
                f"テンプレートから判定した '{self.template_lang}' を使います"
            )
            return self.template_lang
        return code

    def apply_language(self) -> None:
        """全パートの run に校正言語と noProof を書き込む。

        PowerPoint の「校閲 > 言語 > 校正言語の設定」で対象を「ドキュメント」に
        して設定するのと同じ状態にする。
        """
        lang = self.resolve_language()
        no_proof = opt_bool(self.cfg["options"].get("no_proof", True))
        if lang is None and no_proof is None:
            return                           # どちらも「指定しない」
        for root in self._text_roots():
            # <a:rPr> は <a:t> より前に置く必要がある。持たない run には作る
            # （空 run と、Painter を通らないノートのテキストが該当する）。
            for run in list(root.iter(qn("a:r"))):
                if run.find(qn("a:rPr")) is None:
                    run.insert(0, run.makeelement(qn("a:rPr"), {}))
            for el in root.iter():
                if el.tag not in LANG_TAGS:
                    continue
                if lang is not None:
                    el.set("lang", lang)
                if no_proof is not None:
                    el.set("noProof", "1" if no_proof else "0")
        self.applied_language = lang

    def _text_roots(self) -> Iterable[Any]:
        """テキストを持ちうる ppt パートのルート要素。

        slide.notes_slide / prs.notes_master はプロパティを読むだけでパートを
        新規作成してしまうので使わない。パッケージのパートを直接列挙すれば、
        スライド・ノート・ノートマスター・全マスター・全レイアウトを
        副作用なしで網羅できる。
        """
        for part in self.prs.part.package.iter_parts():
            el = getattr(part, "_element", None)   # 画像などは持たない
            if el is not None and str(part.partname).startswith("/ppt/"):
                yield el

    def save(self, out: Path) -> None:
        self.prs.save(out)


# ════════════════════════════════════════════════════════════════════════
# 6. CLI
# ════════════════════════════════════════════════════════════════════════

def outline(specs: list[SlideSpec]) -> str:
    lines = []
    for i, s in enumerate(specs, 1):
        kinds = {"cover": "表紙", "section": "章扉", "content": "本文"}
        body = ", ".join(type(b).__name__ for b in s.blocks) or "-"
        lines.append(f"{i:>3}. [{kinds[s.kind]}] {runs_text(s.title) or '(無題)'}  <{body}>")
    return "\n".join(lines)


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Markdown を PowerPoint に変換します")
    ap.add_argument("input", help="入力 Markdown ファイル")
    ap.add_argument("-o", "--output", help="出力 .pptx（既定: 入力と同名）")
    ap.add_argument("-c", "--config", help="設定 YAML")
    ap.add_argument("-t", "--template", help="テンプレート .pptx/.potx（設定より優先）")
    ap.add_argument("--outline", action="store_true", help="スライド構成だけ表示して終了")
    args = ap.parse_args(argv)

    src_path = Path(args.input)
    if not src_path.exists():
        print(f"入力が見つかりません: {src_path}", file=sys.stderr)
        return 1

    cfg = load_config(args.config)
    if args.template:
        cfg["template"] = args.template

    text = src_path.read_text(encoding="utf-8")
    meta, blocks, notes = parse_markdown(text)
    specs = build_slides(meta, blocks, cfg)
    if notes and specs:
        specs[0].notes.extend(notes[:1])

    if args.outline:
        print(outline(specs))
        return 0

    config_dir = Path(args.config).parent if args.config else None
    renderer = Renderer(cfg, src_path.parent, config_dir)
    renderer.render(specs)

    out = Path(args.output) if args.output else src_path.with_suffix(".pptx")
    renderer.save(out)

    for w in dict.fromkeys(renderer.warnings):
        print(f"warning: {w}", file=sys.stderr)
    print(f"生成しました: {out}  （見出し {len(specs)} → スライド {len(renderer.prs.slides)}）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
