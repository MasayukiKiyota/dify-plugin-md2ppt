"""Tool 実装の結合テスト。_invoke を実際に回す。

dify_plugin SDK が必要なため、SDK を入れた環境で実行する:
    .venv/Scripts/python.exe tests/test_tools.py
"""

from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

PLUGIN_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PLUGIN_DIR))

from dify_plugin.entities.tool import ToolInvokeMessage, ToolRuntime  # noqa: E402

from tools.inspect_template import InspectTemplateTool  # noqa: E402
from tools.md_to_pptx import MdToPptxTool  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"
SAMPLE_MD = (FIXTURES / "sample.md").read_text(encoding="utf-8")
TEMPLATE = (FIXTURES / "template.pptx").read_bytes()
CONFIG = (FIXTURES / "config.yaml").read_text(encoding="utf-8")

failures: list[str] = []


def check(label: str, cond: bool, detail: str = "") -> None:
    if cond:
        print(f"  ok   {label}")
    else:
        print(f"  FAIL {label} {detail}")
        failures.append(label)


class FakeFile:
    """dify_plugin の File のダックタイプ。"""

    def __init__(self, blob: bytes, filename: str, mime_type: str = ""):
        self.blob = blob
        self.filename = filename
        self.mime_type = mime_type


def run(tool_cls, parameters: dict):
    tool = tool_cls(
        runtime=ToolRuntime(credentials={}, user_id="tester", session_id="s1"),
        session=SimpleNamespace(storage=None),
    )
    return list(tool._invoke(parameters))


def by_type(messages, message_type):
    return [m for m in messages if m.type == message_type]


TEXT = ToolInvokeMessage.MessageType.TEXT
JSON = ToolInvokeMessage.MessageType.JSON
BLOB = ToolInvokeMessage.MessageType.BLOB


def payload(msg):
    return msg.message.json_object


def main() -> int:
    tpl_file = FakeFile(TEMPLATE, "template.pptx")

    print("[1] md_to_pptx: markdown_text + template")
    msgs = run(MdToPptxTool, {
        "markdown_text": SAMPLE_MD,
        "template_file": tpl_file,
        "config_yaml": CONFIG,
        "file_name": "quarterly-review",
    })
    texts, jsons, blobs = by_type(msgs, TEXT), by_type(msgs, JSON), by_type(msgs, BLOB)
    check("one blob returned", len(blobs) == 1, str(len(blobs)))
    check("blob emitted last", msgs[-1].type == BLOB, str(msgs[-1].type))
    check("blob is a pptx zip", blobs[0].message.blob[:2] == b"PK")
    check("blob meta has both name keys",
          blobs[0].meta.get("file_name") == "quarterly-review.pptx"
          and blobs[0].meta.get("filename") == "quarterly-review.pptx",
          str(blobs[0].meta))
    check("blob meta mime type",
          blobs[0].meta.get("mime_type", "").endswith("presentationml.presentation"),
          str(blobs[0].meta.get("mime_type")))
    j = payload(jsons[0])
    check("success true", j["success"] is True)
    check("extension appended", j["file_name"] == "quarterly-review.pptx", j["file_name"])
    check("slide count reported", j["slide_count"] == 13, str(j["slide_count"]))
    check("no warnings", j["warnings"] == [], str(j["warnings"]))
    check("front matter in meta", j["meta"].get("author") == "情報システム部",
          str(j["meta"]))
    check("size matches blob", j["size"] == len(blobs[0].message.blob))
    check("text summary present", len(texts) == 1 and "13 枚" in texts[0].message.text,
          texts[0].message.text if texts else "(none)")

    print("[2] md_to_pptx: markdown_file wins over markdown_text, and warns")
    msgs = run(MdToPptxTool, {
        "markdown_text": "# 使われないはずのテキスト\n",
        "markdown_file": FakeFile(
            "---\ntitle: ファイル入力\n---\n\n## 中身\n\n- ok\n".encode("utf-8"),
            "in.md",
        ),
        "template_file": tpl_file,
    })
    j = payload(by_type(msgs, JSON)[0])
    check("file content used", j["meta"].get("title") == "ファイル入力", str(j["meta"]))
    check("warns about both inputs",
          any("両方" in w for w in j["warnings"]), str(j["warnings"]))
    check("default file name", j["file_name"] == "presentation.pptx", j["file_name"])

    print("[3] md_to_pptx: missing template is a clean error, no blob")
    msgs = run(MdToPptxTool, {"markdown_text": SAMPLE_MD})
    check("no blob on error", not by_type(msgs, BLOB))
    check("success false", payload(by_type(msgs, JSON)[0])["success"] is False)
    check("text explains", "テンプレート" in by_type(msgs, TEXT)[0].message.text)

    print("[4] md_to_pptx: empty markdown is a clean error")
    msgs = run(MdToPptxTool, {"markdown_text": "  ", "template_file": tpl_file})
    check("no blob on error", not by_type(msgs, BLOB))
    check("success false", payload(by_type(msgs, JSON)[0])["success"] is False)

    print("[5] md_to_pptx: broken template is a clean error")
    msgs = run(MdToPptxTool, {
        "markdown_text": SAMPLE_MD,
        "template_file": FakeFile(b"not a zip at all", "broken.pptx"),
    })
    check("no blob on error", not by_type(msgs, BLOB))
    check("success false", payload(by_type(msgs, JSON)[0])["success"] is False)
    print("       -> " + by_type(msgs, TEXT)[0].message.text)

    print("[6] inspect_template without emit_config")
    msgs = run(InspectTemplateTool, {"template_file": tpl_file})
    check("no blob", not by_type(msgs, BLOB))
    j = payload(by_type(msgs, JSON)[0])
    check("success true", j["success"] is True)
    check("11 layouts", j["layout_count"] == 11, str(j["layout_count"]))
    check("16:9", j["aspect_ratio"] == "16:9", j["aspect_ratio"])
    check("config_yaml is null", j["config_yaml"] is None)
    check("suggested_config is null", j["suggested_config"] is None)
    report = by_type(msgs, TEXT)[0].message.text
    check("report names the template", "template.pptx" in report)
    check("report lists 表紙", "表紙" in report)

    print("[7] inspect_template with emit_config")
    msgs = run(InspectTemplateTool, {"template_file": tpl_file, "emit_config": True})
    blobs = by_type(msgs, BLOB)
    check("config blob returned", len(blobs) == 1, str(len(blobs)))
    check("blob emitted last", msgs[-1].type == BLOB)
    check("blob named config.yaml", blobs[0].meta.get("file_name") == "config.yaml",
          str(blobs[0].meta))
    j = payload(by_type(msgs, JSON)[0])
    check("config_yaml present", isinstance(j["config_yaml"], str) and j["config_yaml"])
    check("guessed layouts",
          j["suggested_config"]["layouts"]["content"] == "本文",
          str(j["suggested_config"]))

    print("[8] emitted config feeds md_to_pptx unchanged")
    emitted = j["config_yaml"]
    msgs = run(MdToPptxTool, {
        "markdown_text": SAMPLE_MD,
        "template_file": tpl_file,
        "config_yaml": emitted,
    })
    j2 = payload(by_type(msgs, JSON)[0])
    check("round trip succeeds", j2["success"] is True)
    check("round trip has no warnings", j2["warnings"] == [], str(j2["warnings"]))
    check("same slide count", j2["slide_count"] == 13, str(j2["slide_count"]))

    print("[8b] config_yaml 無しでも英語テンプレートが崩れない")
    import io

    from pptx import Presentation

    buf = io.BytesIO()
    Presentation().save(buf)          # python-pptx 既定 = 標準的な英語レイアウト
    msgs = run(MdToPptxTool, {
        "markdown_text": SAMPLE_MD,
        "template_file": FakeFile(buf.getvalue(), "default.pptx"),
    })
    j3 = payload(by_type(msgs, JSON)[0])
    check("no config needed", j3["success"] is True and j3["warnings"] == [],
          str(j3["warnings"]))
    check("layouts auto-detected",
          j3["layouts_used"]["content"] == "Title and Content",
          str(j3["layouts_used"]))
    check("no layout fallback bloat", j3["slide_count"] == 13,
          str(j3["slide_count"]))
    check("blob still returned", len(by_type(msgs, BLOB)) == 1)

    print("[8c] Dify が未入力欄に 'None' を入れて送っても通る（報告された不具合の再現）")
    msgs = run(MdToPptxTool, {
        "markdown_text": SAMPLE_MD,
        "markdown_file": "",
        "template_file": tpl_file,
        "config_yaml": "None",      # <- これが変換全体を止めていた
        "file_name": "None",
    })
    j4 = payload(by_type(msgs, JSON)[0])
    check("変換が成功する", j4["success"] is True, str(j4.get("error")))
    check("警告なし", j4["warnings"] == [], str(j4["warnings"]))
    check("None.pptx にならない", j4["file_name"] == "presentation.pptx",
          j4["file_name"])
    check("blob が返る", len(by_type(msgs, BLOB)) == 1)

    print("[8d] 'None' の markdown_text はファイル入力を邪魔しない")
    msgs = run(MdToPptxTool, {
        "markdown_text": "None",
        "markdown_file": FakeFile(
            "---\ntitle: ファイル入力\n---\n\n## 中身\n\n- ok\n".encode("utf-8"),
            "in.md",
        ),
        "template_file": tpl_file,
    })
    j5 = payload(by_type(msgs, JSON)[0])
    check("誤警告が出ない", j5["warnings"] == [], str(j5["warnings"]))
    check("ファイルの内容が使われる", j5["meta"].get("title") == "ファイル入力",
          str(j5["meta"]))

    print("[8e] emit_config に文字列の 'false' が届いても false として扱う")
    msgs = run(InspectTemplateTool, {"template_file": tpl_file,
                                     "emit_config": "false"})
    check("config blob を返さない", not by_type(msgs, BLOB))
    check("config_yaml は null", payload(by_type(msgs, JSON)[0])["config_yaml"] is None)
    msgs = run(InspectTemplateTool, {"template_file": tpl_file,
                                     "emit_config": "true"})
    check("文字列 'true' は有効", len(by_type(msgs, BLOB)) == 1)

    print("[9] inspect_template: missing file is a clean error")
    msgs = run(InspectTemplateTool, {})
    check("no blob", not by_type(msgs, BLOB))
    check("success false", payload(by_type(msgs, JSON)[0])["success"] is False)

    print()
    if failures:
        print(f"FAILED: {len(failures)} check(s): {', '.join(failures)}")
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
