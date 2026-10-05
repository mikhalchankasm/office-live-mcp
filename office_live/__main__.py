"""python -m office_live [serve|install|postinstall|uninstall|setup|clients|doctor|tools|config]"""

import sys


def main(argv=None) -> int:
    from .cli import _utf8_output

    _utf8_output()
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv and getattr(sys, "frozen", False) and sys.stdin is not None and sys.stdin.isatty():
        # office-live-mcp.exe запущен двойным щелчком (агенты запускают его с каналами, а не с консолью): ставим
        argv = ["install", "--pause"]
    cmd = argv[0] if argv else "serve"
    from . import ConfigError

    try:
        if cmd in ("serve", "run"):
            from .app import main as serve

            serve()
            return 0
        from .cli import run

        return run(cmd, argv[1:])
    except ConfigError as exc:
        print(f"[office-live] Ошибка настройки: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
