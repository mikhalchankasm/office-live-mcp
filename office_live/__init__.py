"""Office Live MCP — управление ОТКРЫТЫМИ документами Excel и Word через COM."""

__version__ = "0.4.0"


class ConfigError(ValueError):
    """Неверная настройка окружения: сервер не должен молча расширять права."""
