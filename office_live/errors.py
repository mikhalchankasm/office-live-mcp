"""Ошибки инструментов.

SDK 2.x показывает модели только `Error executing tool <name>` для любого исключения, кроме
ToolError, поэтому все сообщения для агента обязаны идти через ToolError.
"""

from mcp.server.mcpserver.exceptions import ToolError

__all__ = ["ToolError", "AppBusyError"]


class AppBusyError(ToolError):
    """Office не отвечает: открыт диалог или идёт редактирование ячейки/текста."""
