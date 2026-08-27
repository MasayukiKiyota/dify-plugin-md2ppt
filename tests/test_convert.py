"""SDK 非依存のスモークテスト。

    .venv/Scripts/python.exe tests/test_convert.py

dify_plugin をインストールしなくても変換パイプライン全体を検証できる。
"""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import md2ppt_utils as u  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"
SAMPLE_MD = FIXTURES / "sample.md"
TEMPLATE = FIXTURES / "template.pptx"
CONFIG = FIXTURES / "config.yaml"

_POTX_CT = ("application/vnd.openxmlformats-officedocument"
            ".presentationml.template.main+xml")
_PPTX_CT = ("application/vnd.openxmlformats-officedocument"
            ".presentationml.presentation.main+xml")

failures: list[str] = []


def check(label: str, cond: bool, detail: str = "") -> None:
    if cond:
        print(f"  ok   {label}")
    else:
        print(f"  FAIL {label} {detail}")
        failures.append(label)


def as_potx(pptx_bytes: bytes) -> bytes:
    """.pptx を .potx 相当（content type だけ template）に作り替える。"""
    import io
    src = io.BytesIO(pptx_bytes)
    out = io.BytesIO()
    with zipfile.ZipFile(src) as zf:
        entries = [(i, zf.read(i.filename)) for i in zf.infolist()]
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as w:
        for info, data in entries:
            if info.filename == "[Content_Types].xml":
                data = data.decode("utf-8").replace(_PPTX_CT, _POTX_CT).encode("utf-8")
            w.writestr(info, data)
    return out.getvalue()


def main() -> int:
    md = SAMPLE_MD.read_text(encoding="utf-8")
    tpl = TEMPLATE.read_bytes()
    cfg = CONFIG.read_text(encoding="utf-8")

    print("[1] convert with sample config")
    r = u.convert(md, tpl, "template.pptx", cfg)
    check("returns a zip", r["pptx"][:2] == b"PK")
    check("slide_count >= spec_count",
          r["slide_count"] >= r["spec_count"],
          f"({r['slide_count']} vs {r['spec_count']})")
    check("no warnings", not r["warnings"], str(r["warnings"]))
    check("front matter title picked up", r["meta"].get("title") is not None,
          str(r["meta"]))
    check("outline is non-empty", bool(r["outline"].strip()))
    print(f"       slides={r['slide_count']} specs={r['spec_count']} "
          f"bytes={len(r['pptx'])} meta={r['meta']}")
    print("       " + r["outline"].replace("\n", "\n       "))

    print("[2] convert with DEFAULT_CONFIG (config_yaml 未指定)")
    r2 = u.convert(md, tpl, "template.pptx", None)
    check("same slide count as sample config",
          r2["slide_count"] == r["slide_count"],
          f"({r2['slide_count']} vs {r['slide_count']})")
    check("no warnings", not r2["warnings"], str(r2["warnings"]))

    print("[3] .potx (content type = template) is normalized")
    potx = as_potx(tpl)
    r3 = u.convert(md, potx, "corporate.potx", cfg)
    check("potx converts", r3["pptx"][:2] == b"PK")
    check("potx slide count matches pptx",
          r3["slide_count"] == r["slide_count"])

    print("[4] analyze_template")
    info = u.analyze_template(tpl, "template.pptx")
    check("11 layouts", info["layout_count"] == 11, str(info["layout_count"]))
    check("16:9", info["aspect_ratio"] == "16:9", info["aspect_ratio"])
    check("13.33 x 7.5 in",
          (info["slide_width_in"], info["slide_height_in"]) == (13.33, 7.5),
          f"{info['slide_width_in']} x {info['slide_height_in']}")
    check("表紙 layout present", "表紙" in info["layout_names"],
          str(info["layout_names"]))
    report = u.format_template_report(info)
    check("report mentions every layout",
          all(n in report for n in info["layout_names"]))

    print("[5] suggest_config round-trips through convert")
    sug = u.suggest_config(info)
    check("guessed 表紙/章扉/本文",
          (sug["layouts"]["title"], sug["layouts"]["section"],
           sug["layouts"]["content"]) == ("表紙", "章扉", "本文"),
          str(sug["layouts"]))
    yml = u.config_to_yaml(sug, "template.pptx")
    r5 = u.convert(md, tpl, "template.pptx", yml)
    check("suggested config converts cleanly", not r5["warnings"], str(r5["warnings"]))

    print("[6] analyze_template on .potx")
    info_p = u.analyze_template(potx, "corporate.potx")
    check("potx analyzed", info_p["layout_count"] == 11)

    print("[7] error paths")
    for label, fn in [
        ("bad extension", lambda: u.convert(md, tpl, "deck.ppt", None)),
        ("not a zip", lambda: u.convert(md, b"hello world", "t.pptx", None)),
        ("empty markdown", lambda: u.convert("   ", tpl, "t.pptx", None)),
        ("bad yaml", lambda: u.convert(md, tpl, "t.pptx", "layouts: [oops")),
        ("scalar yaml", lambda: u.convert(md, tpl, "t.pptx", "just a string")),
    ]:
        try:
            fn()
            check(f"{label} raises Md2pptError", False, "(no exception)")
        except u.Md2pptError as e:
            check(f"{label} raises Md2pptError", True)
            print(f"       -> {e}")
        except Exception as e:
            check(f"{label} raises Md2pptError", False,
                  f"(got {type(e).__name__}: {e})")

    print("[7b] 見出しの無い Markdown は本文だけの 1 枚になる")
    r7 = u.convert("plain text only", tpl, "t.pptx", None)
    check("one untitled slide", r7["spec_count"] == 1, str(r7["spec_count"]))

    print("[8] unknown layout names produce warnings, not a crash")
    r8 = u.convert(md, tpl, "t.pptx", 'layouts:\n  content: "存在しないレイアウト"\n')
    check("still converts", r8["pptx"][:2] == b"PK")
    check("warns about the missing layout", any("存在しないレイアウト" in w
                                                for w in r8["warnings"]),
          str(r8["warnings"]))

    print("[9] input helpers")
    check("ensure_pptx appends suffix", u.ensure_pptx("deck") == "deck.pptx")
    check("ensure_pptx sanitizes", u.ensure_pptx("a/b:c.pptx") == "a_b_c.pptx",
          u.ensure_pptx("a/b:c.pptx"))
    check("ensure_pptx falls back", u.ensure_pptx("  ") == "presentation.pptx")
    check("decode cp932", u.decode_text("日本語".encode("cp932")) == "日本語")
    check("decode utf-8 BOM", u.decode_text("# あ".encode("utf-8-sig")) == "# あ")
    text, used = u.read_markdown("# a\r\nb\r\n", None)
    check("CRLF normalized", text == "# a\nb\n", repr(text))
    check("used_file False for text", used is False)

    class FakeFile:
        blob = "# ファイル入力\n".encode("utf-8")
        filename = "in.md"

    text, used = u.read_markdown(None, FakeFile())
    check("File param decoded", text.startswith("# ファイル入力"))
    check("used_file True for file", used is True)
    check("file_name_of", u.file_name_of(FakeFile()) == "in.md")
    check("file_bytes(None) is None", u.file_bytes(None) is None)

    print()
    if failures:
        print(f"FAILED: {len(failures)} check(s): {', '.join(failures)}")
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
