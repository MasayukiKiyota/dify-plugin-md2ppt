"""Dify 向けの薄いラッパ。

md2pptx_core は sample/md2pptx.py の無改造コピーで、パス入出力と `SystemExit` を
前提にした CLI 向けの作りになっている。ここではその差分だけを吸収して、
バイト列 in / バイト列 out のインタフェースを提供する。

このモジュールは dify_plugin を import しない（アップロードファイルは
`.blob` / `.filename` をダックタイピングで読む）。SDK 無しで単体テストできる。
"""

from __future__ import annotations

import copy
import io
import re
import tempfile
import zipfile
from pathlib import Path
from typing import Any

import yaml
from pptx import Presentation

import md2pptx_core as core

EMU_IN = 914400.0

TEMPLATE_SUFFIXES = (".pptx", ".potx")

PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
YAML_MIME = "application/x-yaml"

# .potx は [Content_Types].xml で template.main+xml を宣言する。python-pptx の
# Presentation() は presentation.main+xml しか受け付けないため、拡張子を .pptx に
# 変えるだけでは開けない（md2pptx_core._open_template のやり方では不十分）。
_POTX_CT = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
_PPTX_CT = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"


class Md2pptError(Exception):
    """利用者にそのまま見せられるエラー。"""


# ════════════════════════════════════════════════════════════════════════
# 入力パラメータの取り出し（SDK 非依存）
# ════════════════════════════════════════════════════════════════════════

def file_bytes(param: Any) -> bytes | None:
    """dify_plugin の File / dict / bytes からファイル内容を取り出す。"""
    if param is None or param == "":
        return None
    if isinstance(param, (bytes, bytearray)):
        return bytes(param)
    if isinstance(param, (list, tuple)):
        raise Md2pptError("単一ファイルのパラメータに複数のファイルが渡されました。")

    blob = getattr(param, "blob", None)
    if blob is None and isinstance(param, dict):
        blob = param.get("blob")
    if blob is None:
        raise Md2pptError(
            "アップロードされたファイルの内容を取得できませんでした。"
            "リモートデバッグ中の場合は Dify 側の FILES_URL が"
            "このプラグインから到達可能な URL になっているか確認してください。"
        )
    return bytes(blob)


def file_name_of(param: Any, default: str = "") -> str:
    """File / dict からファイル名を取り出す。"""
    name = getattr(param, "filename", None)
    if not name and isinstance(param, dict):
        name = param.get("filename") or param.get("file_name")
    return str(name or default)


def decode_text(data: bytes) -> str:
    """UTF-8 を基本に、日本語環境でありがちな符号化にフォールバックする。"""
    for enc in ("utf-8-sig", "utf-8", "cp932", "utf-16"):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    raise Md2pptError(
        "ファイルの文字コードを判別できませんでした。UTF-8 で保存し直してください。"
    )


def read_markdown(text_param: Any, file_param: Any) -> tuple[str, bool]:
    """Markdown 本文と、ファイル入力を採用したかどうかを返す。"""
    data = file_bytes(file_param)
    if data is not None:
        text = decode_text(data)
        used_file = True
    else:
        text = text_param if isinstance(text_param, str) else (str(text_param or ""))
        used_file = False
    if not text.strip():
        raise Md2pptError(
            "Markdown が空です。markdown_text か markdown_file のどちらかを指定してください。"
        )
    return text.replace("\r\n", "\n").replace("\r", "\n"), used_file


def ensure_pptx(name: str | None) -> str:
    """出力ファイル名を安全な .pptx 名に整える。"""
    cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", (name or "").strip()).strip(". ")
    if not cleaned:
        cleaned = "presentation"
    return cleaned if cleaned.lower().endswith(".pptx") else f"{cleaned}.pptx"


# ════════════════════════════════════════════════════════════════════════
# 設定
# ════════════════════════════════════════════════════════════════════════

def build_config(config_yaml: str | None) -> dict:
    """YAML 文字列を DEFAULT_CONFIG にマージする。core.load_config のパス版。"""
    if not config_yaml or not config_yaml.strip():
        return copy.deepcopy(core.DEFAULT_CONFIG)
    try:
        data = yaml.safe_load(config_yaml)
    except yaml.YAMLError as e:
        raise Md2pptError(f"config_yaml を解釈できません: {e}") from None
    if data is None:
        return copy.deepcopy(core.DEFAULT_CONFIG)
    if not isinstance(data, dict):
        raise Md2pptError("config_yaml はマッピング（key: value）の形式で記述してください。")
    return core.deep_merge(core.DEFAULT_CONFIG, data)


# ════════════════════════════════════════════════════════════════════════
# テンプレート
# ════════════════════════════════════════════════════════════════════════

def check_template_name(filename: str | None) -> str:
    """拡張子を検証して小文字で返す。"""
    suffix = Path(filename or "template.pptx").suffix.lower()
    if suffix not in TEMPLATE_SUFFIXES:
        raise Md2pptError(
            f"テンプレートは .pptx または .potx を指定してください"
            f"（受け取った拡張子: {suffix or 'なし'}）"
        )
    return suffix


def prepare_template(template_bytes: bytes, workdir: Path) -> Path:
    """テンプレートを workdir に .pptx として書き出し、絶対パスを返す。

    - 常に `.pptx` 名で置くので core._open_template の .potx 分岐
      （後始末されない tempfile.mkdtemp を使う）を通らない。
    - .potx の content type はここで .pptx 相当に書き換える。
    - 絶対パスなので core._find_template が即 return し、Path.cwd() を見ない。
    """
    if not template_bytes:
        raise Md2pptError("テンプレートファイルが空です。")
    if template_bytes[:2] != b"PK":
        raise Md2pptError(
            "テンプレートが Office Open XML（.pptx / .potx）ではありません。"
            "古い .ppt 形式の場合は PowerPoint で .pptx に保存し直してください。"
        )
    dst = workdir / "template.pptx"
    dst.write_bytes(template_bytes)
    _normalize_content_type(dst)
    return dst.resolve()


def _normalize_content_type(path: Path) -> None:
    """.potx の content type を .pptx 相当に書き換える。.pptx なら何もしない。"""
    try:
        with zipfile.ZipFile(path) as zf:
            try:
                ct = zf.read("[Content_Types].xml").decode("utf-8")
            except KeyError:
                raise Md2pptError(
                    "テンプレートの構造が不正です（[Content_Types].xml がありません）。"
                ) from None
            if _POTX_CT not in ct:
                return
            entries = [(info, zf.read(info.filename)) for info in zf.infolist()]
    except zipfile.BadZipFile:
        raise Md2pptError("テンプレートを展開できません。ファイルが破損しています。") from None

    tmp = path.with_name(path.name + ".fixed")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in entries:
            if info.filename == "[Content_Types].xml":
                data = ct.replace(_POTX_CT, _PPTX_CT).encode("utf-8")
            out.writestr(info, data)
    tmp.replace(path)


def _open_presentation(path: Path) -> Presentation:
    try:
        return Presentation(path)
    except Md2pptError:
        raise
    except Exception as e:
        raise Md2pptError(
            f"テンプレートを開けませんでした。破損しているか "
            f"PowerPoint 形式ではない可能性があります: {e}"
        ) from None


# ════════════════════════════════════════════════════════════════════════
# 変換
# ════════════════════════════════════════════════════════════════════════

def convert(
    md_text: str,
    template_bytes: bytes,
    template_filename: str | None = None,
    config_yaml: str | None = None,
) -> dict[str, Any]:
    """Markdown を PowerPoint に変換する。

    returns: {"pptx", "slide_count", "spec_count", "outline", "warnings", "meta"}
    """
    check_template_name(template_filename)
    cfg = build_config(config_yaml)

    with tempfile.TemporaryDirectory(prefix="md2ppt-") as td:
        work = Path(td)
        cfg["template"] = str(prepare_template(template_bytes, work))

        meta, blocks, notes = core.parse_markdown(md_text)
        specs = core.build_slides(meta, blocks, cfg)
        if not specs:
            raise Md2pptError(
                "スライドを 1 枚も作れませんでした。"
                "見出し（# / ##）か front matter の title を含めてください。"
            )
        specs[0].notes.extend(notes[:1])   # core.main() と同じ挙動

        try:
            # base_dir は一時ディレクトリ。相対パスの画像はここに解決される。
            renderer = core.Renderer(cfg, work, work)
            renderer.render(specs)
        except Md2pptError:
            raise
        except SystemExit as e:
            # core._open_template はテンプレート不在時に SystemExit を投げる。
            # BaseException なので except Exception では捕まらない。
            raise Md2pptError(str(e) or "テンプレートの読み込みに失敗しました") from None
        except Exception as e:
            raise Md2pptError(f"PowerPoint の生成に失敗しました: {e}") from None

        buf = io.BytesIO()
        renderer.prs.save(buf)   # python-pptx はファイルライクを受け付ける

        return {
            "pptx": buf.getvalue(),
            "slide_count": len(renderer.prs.slides),
            "spec_count": len(specs),
            "outline": core.outline(specs),
            "warnings": list(dict.fromkeys(renderer.warnings)),
            "meta": {
                k: str(v) for k, v in (meta or {}).items()
                if k in ("title", "subtitle", "author", "date") and v
            },
        }


# ════════════════════════════════════════════════════════════════════════
# テンプレート解析（sample/inspect_template.py 相当）
# ════════════════════════════════════════════════════════════════════════

def _inch(v) -> float | None:
    return None if v is None else round(v / EMU_IN, 2)


def analyze_template(
    template_bytes: bytes, template_filename: str | None = None
) -> dict[str, Any]:
    """レイアウトとプレースホルダの一覧を辞書で返す。"""
    check_template_name(template_filename)
    name = Path(template_filename or "template.pptx").name

    with tempfile.TemporaryDirectory(prefix="md2ppt-") as td:
        path = prepare_template(template_bytes, Path(td))
        prs = _open_presentation(path)

        layouts = []
        for i, layout in enumerate(prs.slide_layouts):
            placeholders = []
            for ph in layout.placeholders:
                pf = ph.placeholder_format
                placeholders.append({
                    "idx": pf.idx,
                    "type": str(pf.type).split(" ")[0],
                    "name": ph.name,
                    "left_in": _inch(ph.left),
                    "top_in": _inch(ph.top),
                    "width_in": _inch(ph.width),
                    "height_in": _inch(ph.height),
                })
            layouts.append({
                "index": i,
                "name": layout.name,
                "placeholders": placeholders,
                "other_shapes": [s.name for s in layout.shapes if not s.is_placeholder],
            })

        w_in = _inch(prs.slide_width) or 0.0
        h_in = _inch(prs.slide_height) or 0.0
        return {
            "template_name": name,
            "slide_width_in": w_in,
            "slide_height_in": h_in,
            "aspect_ratio": _aspect(w_in, h_in),
            "master_count": len(prs.slide_masters),
            "layout_count": len(prs.slide_layouts),
            "layout_names": [l["name"] for l in layouts],
            "layouts": layouts,
        }


def _aspect(w: float, h: float) -> str:
    if not w or not h:
        return "-"
    r = w / h
    if abs(r - 16 / 9) < 0.02:
        return "16:9"
    if abs(r - 4 / 3) < 0.02:
        return "4:3"
    return f"{r:.2f}:1"


def format_template_report(info: dict[str, Any]) -> str:
    """inspect_template.describe() と同じ体裁の人間可読テキスト。"""

    def num(v) -> str:
        return "-" if v is None else f"{v:.2f}"

    out = [
        f"# {info['template_name']}",
        f"スライドサイズ: {num(info['slide_width_in'])}in x "
        f"{num(info['slide_height_in'])}in ({info['aspect_ratio']})",
        f"スライドマスタ数: {info['master_count']} / レイアウト数: {info['layout_count']}",
        "",
    ]
    for layout in info["layouts"]:
        out.append(f"[{layout['index']}] {layout['name']!r}")
        out.append(
            "      idx  type                 name                      "
            "left   top    width  height  (inch)"
        )
        for ph in layout["placeholders"]:
            out.append(
                f"      {ph['idx']:<4} {ph['type']:<20} {ph['name'][:24]:<25} "
                f"{num(ph['left_in']):>6} {num(ph['top_in']):>6} "
                f"{num(ph['width_in']):>6} {num(ph['height_in']):>6}"
            )
        if layout["other_shapes"]:
            out.append(
                f"      (プレースホルダ以外の図形: {', '.join(layout['other_shapes'])})"
            )
        out.append("")
    return "\n".join(out)


_GUESS = {
    "title": ("表紙", "title slide", "タイトル スライド"),
    "section": ("章扉", "section", "セクション"),
    "content": ("本文", "title and content", "タイトルとコンテンツ"),
    "table": ("タイトルのみ", "title only"),
    "blank": ("白紙", "blank"),
}


def suggest_config(info: dict[str, Any]) -> dict[str, Any]:
    """レイアウト名から config の推測値を組み立てる。emit_config() 相当。"""
    names = info["layout_names"]

    def guess(keys: tuple[str, ...]) -> str:
        for name in names:
            low = name.lower()
            if any(k.lower() in low for k in keys):
                return name
        return names[0] if names else ""

    layouts = {key: guess(keys) for key, keys in _GUESS.items()}

    # content レイアウトの本文プレースホルダ idx を実測する（既定の 1 とは限らない）
    body_idx = 1
    content = next((l for l in info["layouts"] if l["name"] == layouts["content"]), None)
    if content:
        body = next(
            (p for p in content["placeholders"]
             if p["idx"] != 0 and p["type"] in ("BODY", "OBJECT", "SUBTITLE")),
            None,
        )
        if body:
            body_idx = body["idx"]

    return {
        "layouts": layouts,
        "placeholders": {"title": 0, "subtitle": 1, "body": body_idx},
    }


def config_to_yaml(suggested: dict[str, Any], template_name: str) -> str:
    """suggest_config() の結果を config_yaml パラメータに貼れる YAML にする。"""
    body = yaml.safe_dump(
        suggested, allow_unicode=True, sort_keys=False, default_flow_style=False
    )
    return (
        f"# md2ppt 設定（inspect_template ツールが {template_name} から生成）\n"
        f"# md_to_pptx ツールの config_yaml パラメータにこのまま貼り付けてください。\n"
        f"{body}"
    )
