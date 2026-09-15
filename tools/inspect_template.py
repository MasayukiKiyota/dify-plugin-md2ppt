from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from md2ppt_utils import (
    Md2pptError,
    analyze_template,
    bool_param,
    config_to_yaml,
    detect_config,
    file_bytes,
    file_name_of,
    format_template_report,
    payload_messages,
)


class InspectTemplateTool(Tool):
    def _invoke(
        self, tool_parameters: dict[str, Any]
    ) -> Generator[ToolInvokeMessage, None, None]:
        emit_config = bool_param(tool_parameters.get("emit_config"))

        try:
            template = tool_parameters.get("template_file")
            if template is None or template == "":
                yield self.create_text_message(
                    "解析する PowerPoint テンプレート（.pptx / .potx）を"
                    "アップロードしてください。"
                )
                yield from payload_messages(
                    self, {"success": False, "error": "template_file is required"}
                )
                return

            template_name = file_name_of(template, "template.pptx")
            info = analyze_template(file_bytes(template), template_name)
        except Md2pptError as e:
            yield self.create_text_message(str(e))
            yield from payload_messages(self, {"success": False, "error": str(e)})
            return
        except Exception as e:
            yield self.create_text_message(f"テンプレートの解析に失敗しました: {e}")
            yield from payload_messages(self, {"success": False, "error": str(e)})
            return

        report = format_template_report(info)

        config_yaml = None
        if emit_config:
            # 判定結果は YAML テキストにして返す。ファイルでは返さない
            # （後続のノードからは出力変数 config_yaml をそのまま渡せる）。
            config_yaml = config_to_yaml(detect_config(info), info["template_name"])
            report = f"{report}\n{config_yaml}"

        yield self.create_text_message(report)

        yield from payload_messages(self, {
            "success": True,
            "template_name": info["template_name"],
            "slide_width_in": info["slide_width_in"],
            "slide_height_in": info["slide_height_in"],
            "aspect_ratio": info["aspect_ratio"],
            "master_count": info["master_count"],
            "layout_count": info["layout_count"],
            "layout_names": info["layout_names"],
            "language": info["language"],
            "layouts": info["layouts"],
            "config_yaml": config_yaml,
        })
