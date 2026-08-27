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
    "fonts": {
        "latin": "Calibri",
        "eastasian": "Yu Gothic",
        "code": "Consolas",
        "code_eastasian": "MS Gothic",
    },
    "sizes": {
        "title": 32,
        "section_title": 40,
        "cover_title": 40,
        "cover_subtitle": 18,
        "h3": 20,
        "h4": 17,
        "body": [18, 16, 14, 13, 12],   # 箇条書きレベル別
        "paragraph": 16,
        "table": 12,
        "table_header": 12,
        "code": 12,
        "quote": 15,
        "caption": 11,
        "page_number": 10,
        "min_body": 10,                 # 収まらないときの下限
    },
    "colors": {
        "text": "1F2430",
        "heading": "1E2A44",
        "accent": "1E6F8E",
        "muted": "6B7280",
        "code_text": "24292F",
        "code_bg": "F3F4F6",
        "quote_text": "4A5568",
        "quote_bar": "1E6F8E",
        "table_header_bg": "1E2A44",
        "table_header_text": "FFFFFF",
        "table_band_bg": "F2F4F7",
        "table_text": "1F2430",
        "table_border": "C9D2DC",
    },
    "table": {
        # PowerPoint 組み込みスタイル ID（テンプレート側のスタイルを使うなら null）
        "style_id": "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",  # Medium Style 2 - Accent 1
        "first_row_header": True,
        "banding": True,
        # true にすると colors.table_* で明示的に塗る（どの環境でも同じ見た目）
        # false にするとテンプレート/スタイル側の書式に任せる
        "explicit_format": True,
        "min_row_height": 0.32,   # inch
        "cell_margin": 0.06,      # inch
    },
    "quote": {"bar": True, "indent": 0.3, "italic": True},
    "image": {"max_height": 4.2, "align": "center"},
    "spacing": {
        "block_gap": 8,          # pt
        "para_space_after": 6,   # pt
        "line_ratio": 1.38,
        # 箇条書きのインデント（inch）。null にするとテンプレートの設定を継承
        "list_indent": 0.3,
        "list_hanging": 0.3,
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
    cap = max(width_pt / size_pt, 1.0)
    lines = 0
    for logical in text.split("\n"):
        u = char_units(logical)
        lines += max(1, math.ceil(u / cap))
    return max(lines, 1)


def rgb(hexstr: str) -> RGBColor:
    return RGBColor.from_string(hexstr.replace("#", "").upper())


def set_typeface(rPr, latin: str, ea: str) -> None:
    """<a:latin> の直後に <a:ea>/<a:cs> を差し込む（日本語が明朝に落ちるのを防ぐ）。"""
    latin_el = rPr.find(qn("a:latin"))
    if latin_el is None:
        latin_el = rPr.makeelement(qn("a:latin"), {"typeface": latin})
        rPr.append(latin_el)
    else:
        latin_el.set("typeface", latin)
    idx = list(rPr).index(latin_el)
    for tag, face in ((qn("a:ea"), ea), (qn("a:cs"), latin)):
        el = rPr.find(tag)
        if el is None:
            el = rPr.makeelement(tag, {"typeface": face})
            idx += 1
            rPr.insert(idx, el)
        else:
            el.set("typeface", face)
            idx = list(rPr).index(el)


class Painter:
    """設定に基づいて run / 段落を装飾するヘルパ。"""

    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.f = cfg["fonts"]
        self.c = cfg["colors"]
        self.s = cfg["sizes"]

    def style_run(self, run, *, size: float, bold=False, italic=False,
                  color: str | None = None, mono=False, strike=False,
                  link: str | None = None) -> None:
        font = run.font
        font.size = Pt(size)
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

    def fill_runs(self, para, runs: Sequence[Run], *, size: float,
                  color: str | None = None, base_bold=False,
                  base_italic=False) -> None:
        """Run 列を 1 段落に流し込む（改行 Run はソフト改行として扱う）。"""
        color = color or self.c["text"]
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
                size=size * (0.94 if r.code else 1.0),
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
        self.base_dir = base_dir          # 画像の相対パス基準（Markdown のあるフォルダ）
        self.paint = Painter(cfg)
        self.prs = self._open_template(self._find_template(cfg["template"], config_dir, base_dir))
        self.layouts = {l.name: l for l in self.prs.slide_layouts}
        self.slide_w_pt = self.prs.slide_width / EMU_PER_PT
        self.slide_h_pt = self.prs.slide_height / EMU_PER_PT
        self.warnings: list[str] = []

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

    def content_rect_pt(self) -> tuple[float, float, float, float]:
        ba = self.cfg["body_area"]
        if all(ba.get(k) is not None for k in ("left", "top", "width", "height")):
            return (ba["left"] * 72, ba["top"] * 72, ba["width"] * 72, ba["height"] * 72)
        layout = self.layout("content")
        body_idx = self.cfg["placeholders"]["body"]
        for ph in layout.placeholders:
            if ph.placeholder_format.idx == body_idx and ph.left is not None:
                return (ph.left / EMU_PER_PT, ph.top / EMU_PER_PT,
                        ph.width / EMU_PER_PT, ph.height / EMU_PER_PT)
        # フォールバック: 上下左右にマージン
        m = 0.7 * 72
        top = 1.6 * 72
        return (m, top, self.slide_w_pt - 2 * m, self.slide_h_pt - top - 0.7 * 72)

    # ---------- 高さ見積り ----------

    def measure(self, blk: Block, width_pt: float, scale: float = 1.0) -> float:
        s, sp = self.cfg["sizes"], self.cfg["spacing"]
        lr = sp["line_ratio"]
        if isinstance(blk, Heading):
            size = (s["h3"] if blk.level <= 3 else s["h4"]) * scale
            return wrapped_lines(runs_text(blk.runs), size, width_pt) * size * lr + 6
        if isinstance(blk, Para):
            size = s["paragraph"] * scale
            return wrapped_lines(runs_text(blk.runs), size, width_pt) * size * lr + sp["para_space_after"]
        if isinstance(blk, ListBlock):
            total = 0.0
            for it in blk.items:
                size = self.body_size(it.level) * scale
                indent = 18 * (it.level + 1)
                total += wrapped_lines(runs_text(it.runs), size, width_pt - indent) * size * lr
                total += sp["para_space_after"] * 0.6
            return total
        if isinstance(blk, Table):
            return self.table_height(blk, width_pt, scale)
        if isinstance(blk, Code):
            size = s["code"] * scale
            lines = sum(max(1, math.ceil(char_units(l) * 0.9 / max(width_pt / size, 1)))
                        for l in blk.text.split("\n")) or 1
            return lines * size * 1.28 + 16
        if isinstance(blk, Quote):
            inner = width_pt - self.cfg["quote"]["indent"] * 72
            return sum(self.measure(b, inner, scale) for b in blk.blocks) + 8
        if isinstance(blk, Image):
            return self.image_size_pt(blk, width_pt)[1] + 6
        if isinstance(blk, Rule):
            return 12
        return 0.0

    def body_size(self, level: int) -> float:
        sizes = self.cfg["sizes"]["body"]
        return sizes[min(level, len(sizes) - 1)]

    def table_height(self, tb: Table, width_pt: float, scale: float = 1.0) -> float:
        size = self.cfg["sizes"]["table"] * scale
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

    def set_title(self, slide, runs: Sequence[Run], size: float,
                  color: str | None = None) -> None:
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
        # タイトルが長い場合は自動で縮小
        while size > 14 and wrapped_lines(text, size, width_pt) * size * 1.2 > height_pt:
            size -= 2
        self.paint.fill_runs(para, runs, size=size,
                             color=color or self.cfg["colors"]["heading"], base_bold=True)
        self.paint.no_bullet(para)

    def render_cover(self, spec: SlideSpec) -> None:
        slide = self.add_slide("title")
        self.set_title(slide, spec.title, self.cfg["sizes"]["cover_title"])
        sub = self._ph(slide, self.cfg["placeholders"]["subtitle"])
        if sub is not None:
            if spec.subtitle:
                tf = sub.text_frame
                tf.word_wrap = True
                para = tf.paragraphs[0]
                self.paint.fill_runs(para, spec.subtitle,
                                     size=self.cfg["sizes"]["cover_subtitle"],
                                     color=self.cfg["colors"]["muted"])
                self.paint.no_bullet(para)
            else:
                self.drop(sub)
        self.clear_empty_placeholders(slide, {self.cfg["placeholders"]["title"]})
        self.attach_notes(slide, spec)

    def render_section(self, spec: SlideSpec) -> None:
        slide = self.add_slide("section")
        self.set_title(slide, spec.title, self.cfg["sizes"]["section_title"])
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
            self.set_title(slide, spec.title, self.cfg["sizes"]["title"])
            self.clear_empty_placeholders(slide, {self.cfg["placeholders"]["title"]})
            self.attach_notes(slide, spec)
            return

        chunks = self.paginate(spec.blocks)
        for n, (blocks, scale) in enumerate(chunks):
            title = list(spec.title)
            if n > 0 and title:
                title = title + [Run(opts["continuation_suffix"])]
            if self.is_text_only(blocks):
                slide = self.render_text_slide(title, blocks, scale)
            else:
                slide = self.render_rich_slide(title, blocks, scale)
            if n == 0:
                self.attach_notes(slide, spec)

    def paginate(self, blocks: list[Block]) -> list[tuple[list[Block], float]]:
        """収まるように分割し、(ブロック列, フォント倍率) を返す。"""
        opts = self.cfg["options"]
        left, top, width, height = self.content_rect_pt()
        gap = self.cfg["spacing"]["block_gap"]

        # まず全体が入るか、段階的に縮小しながら確認
        for step in range(opts["shrink_steps"] + 1):
            scale = 1.0 - 0.08 * step
            total = sum(self.measure(b, width, scale) for b in blocks) + gap * (len(blocks) - 1)
            if total <= height:
                return [(blocks, scale)]
        if not opts["auto_split"]:
            return [(blocks, 1.0 - 0.08 * opts["shrink_steps"])]

        scale = 1.0 - 0.08 * opts["shrink_steps"]
        chunks: list[tuple[list[Block], float]] = []
        cur: list[Block] = []
        used = 0.0
        for blk in blocks:
            pieces = self.split_block(blk, width, height, scale)
            for piece in pieces:
                h = self.measure(piece, width, scale)
                if cur and used + gap + h > height:
                    chunks.append((cur, scale))
                    cur, used = [], 0.0
                cur.append(piece)
                used += h + gap
        if cur:
            chunks.append((cur, scale))
        return chunks or [(blocks, scale)]

    def split_block(self, blk: Block, width: float, height: float, scale: float) -> list[Block]:
        """単体でスライドに収まらないブロックを分割する（表とリスト）。"""
        if self.measure(blk, width, scale) <= height:
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
                if cur.items and self.measure(trial, width, scale) > height:
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
        self.set_title(slide, title, self.cfg["sizes"]["title"])
        body = self._ph(slide, self.cfg["placeholders"]["body"])
        if body is None:
            self.clear_empty_placeholders(slide, {self.cfg["placeholders"]["title"]})
            left, top, width, height = self.content_rect_pt()
            self.flow_blocks(slide, blocks, left, top, width, height, scale)
            return slide
        tf = body.text_frame
        tf.word_wrap = True
        self.fill_text_frame(tf, blocks, scale)
        self.clear_empty_placeholders(
            slide, {self.cfg["placeholders"]["title"], self.cfg["placeholders"]["body"]})
        return slide

    def render_rich_slide(self, title: list[Run], blocks: list[Block], scale: float):
        """表・コード・画像を含むスライド → 手動レイアウト。"""
        slide = self.add_slide("table")
        self.set_title(slide, title, self.cfg["sizes"]["title"])
        self.clear_empty_placeholders(slide, {self.cfg["placeholders"]["title"]})
        left, top, width, height = self.content_rect_pt()
        self.flow_blocks(slide, blocks, left, top, width, height, scale)
        return slide

    # --- テキストフレームへの流し込み ---

    def fill_text_frame(self, tf, blocks: Sequence[Block], scale: float,
                        first_para=None) -> None:
        p = self.cfg
        sizes = p["sizes"]
        started = False

        def new_para():
            nonlocal started
            if not started:
                started = True
                return tf.paragraphs[0]
            return tf.add_paragraph()

        for blk in blocks:
            if isinstance(blk, Heading):
                para = new_para()
                size = (sizes["h3"] if blk.level <= 3 else sizes["h4"]) * scale
                self.paint.fill_runs(para, blk.runs, size=size,
                                     color=p["colors"]["accent"], base_bold=True)
                self.paint.no_bullet(para)
                self.paint.indent(para, 0, hanging=False)
                para.space_before = Pt(8 if started else 0)
                para.space_after = Pt(3)
            elif isinstance(blk, Para):
                para = new_para()
                self.paint.fill_runs(para, blk.runs, size=sizes["paragraph"] * scale,
                                     color=p["colors"]["text"])
                self.paint.no_bullet(para)
                self.paint.indent(para, 0, hanging=False)
                para.space_after = Pt(p["spacing"]["para_space_after"])
            elif isinstance(blk, ListBlock):
                for it in blk.items:
                    para = new_para()
                    para.level = min(it.level, 4)
                    self.paint.fill_runs(para, it.runs,
                                         size=self.body_size(it.level) * scale,
                                         color=p["colors"]["text"])
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
                    self.paint.fill_runs(para, runs, size=sizes["quote"] * scale,
                                         color=p["colors"]["quote_text"],
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
            h = self.measure(blk, width, scale)
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
        line.line.color.rgb = rgb(self.cfg["colors"]["table_border"])
        line.line.width = Pt(1)
        line.shadow.inherit = False

    def draw_quote(self, slide, q: Quote, x: float, y: float, w: float,
                   h: float, scale: float) -> None:
        indent = self.cfg["quote"]["indent"] * 72
        if self.cfg["quote"]["bar"]:
            from pptx.enum.shapes import MSO_SHAPE
            bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(3), Pt(h))
            bar.fill.solid()
            bar.fill.fore_color.rgb = rgb(self.cfg["colors"]["quote_bar"])
            bar.line.fill.background()
            bar.shadow.inherit = False
        box = self.textbox(slide, x + indent, y, w - indent, h)
        self.fill_text_frame(box.text_frame, q.blocks, scale)
        for para in box.text_frame.paragraphs:
            for run in para.runs:
                run.font.italic = self.cfg["quote"]["italic"]
                run.font.color.rgb = rgb(self.cfg["colors"]["quote_text"])

    def draw_code(self, slide, code: Code, x: float, y: float, w: float,
                  h: float, scale: float) -> None:
        from pptx.enum.shapes import MSO_SHAPE
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h))
        shape.adjustments[0] = 0.04
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(self.cfg["colors"]["code_bg"])
        shape.line.color.rgb = rgb(self.cfg["colors"]["table_border"])
        shape.line.width = Pt(0.75)
        shape.shadow.inherit = False
        tf = shape.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = Pt(9)
        tf.margin_top = tf.margin_bottom = Pt(7)
        tf.vertical_anchor = MSO_ANCHOR.TOP
        size = self.cfg["sizes"]["code"] * scale
        lines = code.text.split("\n")
        for i, line in enumerate(lines):
            para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            para.alignment = PP_ALIGN.LEFT
            self.paint.no_bullet(para)
            run = para.add_run()
            run.text = line if line else " "
            self.paint.style_run(run, size=size, color=self.cfg["colors"]["code_text"], mono=True)
            para.line_spacing = 1.0
            para.space_after = Pt(0)

    def draw_image(self, slide, img: Image, x: float, y: float, w: float, scale: float) -> None:
        path = self.resolve_image(img.src)
        if path is None:
            box = self.textbox(slide, x, y, w, 20)
            para = box.text_frame.paragraphs[0]
            self.paint.fill_runs(para, [Run(f"[画像が見つかりません: {img.src}]")],
                                 size=self.cfg["sizes"]["caption"],
                                 color=self.cfg["colors"]["muted"])
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

        size = self.cfg["sizes"]["table"] * scale
        head_size = self.cfg["sizes"]["table_header"] * scale
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
                    color=self.cfg["colors"]["table_header_text"] if is_head
                    else self.cfg["colors"]["table_text"],
                    base_bold=is_head,
                )
                align = tb.aligns[ci] if ci < len(tb.aligns) else "left"
                para.alignment = {"center": PP_ALIGN.CENTER,
                                  "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)

    def fill_cell(self, cell, ri: int, is_head: bool) -> None:
        c = self.cfg["colors"]
        cell.fill.solid()
        if is_head:
            cell.fill.fore_color.rgb = rgb(c["table_header_bg"])
        elif self.cfg["table"]["banding"] and ri % 2 == 0:
            cell.fill.fore_color.rgb = rgb(c["table_band_bg"])
        else:
            cell.fill.fore_color.rgb = rgb("FFFFFF")

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
        size = self.cfg["sizes"]["page_number"]
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
            self.paint.style_run(run, size=size, color=self.cfg["colors"]["muted"])

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
