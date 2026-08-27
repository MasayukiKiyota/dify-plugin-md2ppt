from collections.abc import Mapping
from typing import Any

from dify_plugin import ToolProvider


class Md2pptProvider(ToolProvider):
    def _validate_credentials(self, credentials: Mapping[str, Any]) -> None:
        """このプラグインは外部通信を行わず認証情報を必要としないため、検証は不要。"""
        return
