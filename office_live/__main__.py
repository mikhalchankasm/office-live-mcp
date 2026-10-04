"""python -m office_live [serve|doctor|tools|config]"""

import sys


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    cmd = argv[0] if argv else "serve"
    if cmd in ("serve", "run"):
        from .app import main as serve

        serve()
        return 0
    from .cli import run

    return run(cmd, argv[1:])


if __name__ == "__main__":
    sys.exit(main())
