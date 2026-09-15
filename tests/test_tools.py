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

import yaml  # noqa: E402
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
VARIABLE = ToolInvokeMessage.MessageType.VARIABLE

# Dify の tool ノードが自前で使う出力キー。変数名がぶつかると上書きされる。
RESERVED = {"text", "files", "json"}


def payload(msg):
    return msg.message.json_object


def variable_names(messages) -> list:
    """出力された変数名。重複を検出したいので list で返す。"""
    return [m.message.variable_name for m in by_type(messages, VARIABLE)]


def variables(messages) -> dict:
    return {m.message.variable_name: m.message.variable_value
            for m in by_type(messages, VARIABLE)}


def message_order_ok(messages) -> bool:
    """変数は JSON より後、blob より前に出ていること。"""
    types = [m.type for m in messages]
    at = [i for i, t in enumerate(types) if t == VARIABLE]
    if not at:
        return False
    if JSON in types and types.index(JSON) > at[0]:
        return False
    return BLOB not in types or at[-1] < types.index(BLOB)


def check_vars(label: str, msgs) -> dict:
    """JSON ペイロードと変数メッセージが 1 対 1 で対応していること。

    output_schema に名前を宣言しただけでは値が入らない（＝後続ノードから
    参照すると null になる）という不具合を構造的に防ぐためのチェック。
    """
    jsons = by_type(msgs, JSON)
    check(f"{label}: JSON は 1 つ", len(jsons) == 1, str(len(jsons)))
    j, names = payload(jsons[0]), variable_names(msgs)
    v = variables(msgs)
    check(f"{label}: 変数名が JSON のキーと一致",
          set(names) == set(j), str(set(names) ^ set(j)))
    check(f"{label}: 変数名が重複しない", len(names) == len(set(names)), str(names))
    check(f"{label}: 予約名と衝突しない", not (RESERVED & set(names)), str(names))
    check(f"{label}: 変数は JSON の後・blob の前",
          message_order_ok(msgs), str([m.type.value for m in msgs]))
    return v


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
    v1 = check_vars("[1]", msgs)
    check("slide_count が変数に入る", v1["slide_count"] == 13, str(v1["slide_count"]))
    check("outline が変数に入る", isinstance(v1["outline"], str) and v1["outline"])
    check("warnings が変数に入る", v1["warnings"] == [], str(v1["warnings"]))
    check("成功時は error を出さない", "error" not in v1, str(sorted(v1)))
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
    check("suggested_config は出力から消えた", "suggested_config" not in j,
          str(sorted(j)))
    v6 = check_vars("[6]", msgs)
    check("config_yaml 変数も null", "config_yaml" in v6 and v6["config_yaml"] is None,
          str(sorted(v6)))
    report = by_type(msgs, TEXT)[0].message.text
    check("report names the template", "template.pptx" in report)
    check("report lists 表紙", "表紙" in report)
    check("判定した校正言語を返す", j["language"] == "en-US", str(j.get("language")))
    check("report に校正言語が出る", "校正言語" in report)

    print("[7] inspect_template with emit_config")
    msgs = run(InspectTemplateTool, {"template_file": tpl_file, "emit_config": True})
    check("ファイルは返さない", not by_type(msgs, BLOB))
    check("最後は変数メッセージ", msgs[-1].type == VARIABLE, str(msgs[-1].type))
    j = payload(by_type(msgs, JSON)[0])
    v7 = check_vars("[7]", msgs)
    check("config_yaml present", isinstance(j["config_yaml"], str) and j["config_yaml"])
    check("config_yaml 変数に本文が入る", v7["config_yaml"] == j["config_yaml"])
    check("生成元のコメントが付く", v7["config_yaml"].startswith("# md2ppt"),
          v7["config_yaml"][:40])
    # config.yaml ファイルが無くなり YAML テキストが唯一の受け渡し経路になったので、
    # 実際に YAML として読めることまで確かめる。
    cfg = yaml.safe_load(j["config_yaml"])
    check("guessed layouts", cfg["layouts"]["content"] == "本文", str(cfg["layouts"]))
    check("校正言語も config に乗る", cfg["options"]["language"] == "en-US",
          str(cfg.get("options")))
    check("emit した YAML に language がある", "language:" in j["config_yaml"])

    print("[8] emitted config feeds md_to_pptx unchanged")
    emitted = v7["config_yaml"]      # 実ワークフローと同じ「変数 → 次ノード」の経路
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
    check("config_yaml は null", variables(msgs)["config_yaml"] is None)
    msgs = run(InspectTemplateTool, {"template_file": tpl_file,
                                     "emit_config": "true"})
    emitted_yaml = variables(msgs)["config_yaml"]
    check("文字列 'true' は有効",
          isinstance(emitted_yaml, str) and "layouts:" in emitted_yaml,
          str(emitted_yaml)[:40])

    print("[9b] プラグインの YAML が全部読める")
    # ここが壊れると dify plugin package が
    #   "mapping values are not allowed in this context" で落ちる。
    # 説明文に ": "（コロン＋空白）を素で書くと平文スカラーが切れるのが典型。
    root = Path(__file__).resolve().parent.parent
    yaml_files = sorted(
        list(root.glob("*.yaml")) + list(root.glob("provider/*.yaml"))
        + list(root.glob("tools/*.yaml"))
    )
    check("YAML が 4 つ以上ある", len(yaml_files) >= 4, str(len(yaml_files)))
    for path in yaml_files:
        rel = path.relative_to(root).as_posix()
        try:
            loaded = yaml.safe_load(path.read_text(encoding="utf-8"))
            check(f"{rel} が読める", isinstance(loaded, dict))
        except yaml.YAMLError as e:                             # noqa: PERF203
            check(f"{rel} が読める", False, str(e).splitlines()[0])

    manifest = yaml.safe_load((root / "manifest.yaml").read_text(encoding="utf-8"))
    for rel in manifest["plugins"]["tools"]:
        spec = yaml.safe_load((root / rel).read_text(encoding="utf-8"))
        check(f"{rel} に tools 定義がある",
              isinstance(spec.get("tools"), list) and bool(spec["tools"]))

    print("[9c] output_schema と実際に出る変数が一致する")
    # output_schema は名前と型を宣言するだけで値は入らない。宣言だけ足して
    # 変数を出し忘れると、後続ノードから参照したときに黙って null になる。
    provider = yaml.safe_load(
        (root / "provider/md2ppt.yaml").read_text(encoding="utf-8"))
    props = {}
    for rel in provider["tools"]:
        spec = yaml.safe_load((root / rel).read_text(encoding="utf-8"))
        props[spec["identity"]["name"]] = set(spec["output_schema"]["properties"])
    check("2 つのツールの output_schema を読めた", len(props) == 2, str(sorted(props)))

    ok_cases = [
        ("md_to_pptx", run(MdToPptxTool, {"markdown_text": SAMPLE_MD,
                                          "template_file": tpl_file})),
        ("inspect_template", run(InspectTemplateTool, {"template_file": tpl_file,
                                                       "emit_config": True})),
    ]
    for name, ok_msgs in ok_cases:
        got = set(variable_names(ok_msgs))
        want = props[name] - {"error"}
        check(f"{name}: 成功時の変数が output_schema を過不足なく満たす",
              got == want, str(sorted(got ^ want)))

    for name, err_msgs in (("md_to_pptx", run(MdToPptxTool, {"markdown_text": SAMPLE_MD})),
                           ("inspect_template", run(InspectTemplateTool, {}))):
        got = set(variable_names(err_msgs))
        check(f"{name}: エラー時は success / error だけ",
              got == {"success", "error"}, str(sorted(got)))
        check(f"{name}: それも output_schema にある", got <= props[name])

    print("[9] inspect_template: missing file is a clean error")
    msgs = run(InspectTemplateTool, {})
    check("no blob", not by_type(msgs, BLOB))
    check("success false", payload(by_type(msgs, JSON)[0])["success"] is False)
    check("エラー時は success / error だけ",
          set(variable_names(msgs)) == {"success", "error"},
          str(sorted(variable_names(msgs))))
    check("success が false として変数に入る", variables(msgs)["success"] is False)

    print()
    if failures:
        print(f"FAILED: {len(failures)} check(s): {', '.join(failures)}")
        return 1
    print("all checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
