"""Сборка сервера: импорт модулей с инструментами регистрирует их в реестре."""

from .registry import CATALOG, mcp  # noqa: F401

# порядок важен только для порядка инструментов в tools/list
from . import bridge, evaltool, excel_analysis, excel_core, excel_format, excel_pivot, journal, navigation, templates, undo, window, word_core, word_edits, word_layout, word_tables  # noqa: F401,E402


def main() -> None:
    """Запуск stdio MCP-сервера."""
    mcp.run()


if __name__ == "__main__":
    main()
