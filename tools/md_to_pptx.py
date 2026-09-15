from collections.abc import Generator
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from md2ppt_utils import (
    PPTX_MIME,
    Md2pptError,
    convert,
    ensure_pptx,
    file_bytes,
    file_name_of,
    payload_messages,
    read_markdown,
    text_param,
)


class MdToPptxTool(Tool):
    def _invoke(
        self, tool_parameters: dict[str, Any]
    ) -> Generator[ToolInvokeMessage, None, None]:
        notices: list[str] = []

        try:
            markdown_file = tool_parameters.get("markdown_file")
            markdown, used_file = read_markdown(
                tool_parameters.get("markdown_text"), markdown_file
            )
            if used_file and text_param(tool_parameters.get("markdown_text")):
                notices.append(
                    "Markdown テキストと Markdown ファイルの両方が指定されたため、"
                    "ファイルの内容を使用しました。"
                )

            template = tool_parameters.get("template_file")
            if template is None or template == "":
                yield self.create_text_message(
                    "PowerPoint テンプレート（.pptx / .potx）をアップロードしてください。"
                )
                yield from payload_messages(
                    self, {"success": False, "error": "template_file is required"}
                )
                return
            template_bytes = file_bytes(template)
            template_name = file_name_of(template, "template.pptx")

            file_name = ensure_pptx(tool_parameters.get("file_name"))

            result = convert(
                markdown,
                template_bytes,
                template_name,
                tool_parameters.get("config_yaml"),
            )
        except Md2pptError as e:
            yield self.create_text_message(str(e))
            yield from payload_messages(self, {"success": False, "error": str(e)})
            return
        except Exception as e:
            yield self.create_text_message(f"変換に失敗しました: {e}")
            yield from payload_messages(self, {"success": False, "error": str(e)})
            return

        pptx: bytes = result["pptx"]
        warnings = notices + result["warnings"]

        summary = (
            f"{result['slide_count']} 枚のスライドを生成しました（{file_name}）。"
        )
        if warnings:
            summary += "\n" + "\n".join(f"警告: {w}" for w in warnings)
        yield self.create_text_message(summary)

        yield from payload_messages(self, {
            "success": True,
            "file_name": file_name,
            "mime_type": PPTX_MIME,
            "size": len(pptx),
            "slide_count": result["slide_count"],
            "spec_count": result["spec_count"],
            "template_name": template_name,
            "meta": result["meta"],
            "layouts_used": result["layouts_used"],
            "language_used": result["language_used"],
            "outline": result["outline"],
            "warnings": warnings,
        })

        # 変数を先に流してから、生成ファイルを最後に emit する（Dify 側では
        # ファイル出力として扱われ、後続からは files[0] で参照する）。
        # file_name / filename は Dify のバージョン差を吸収するため両方入れる。
        yield self.create_blob_message(
            pptx,
            meta={
                "file_name": file_name,
                "filename": file_name,
                "mime_type": PPTX_MIME,
            },
        )
