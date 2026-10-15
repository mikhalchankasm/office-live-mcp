"""Windows numeric release versions, independent of PyInstaller."""

import re


def numeric_version(version: str) -> tuple[int, int, int, int]:
    """Take up to four leading release components, ignoring prerelease/dev suffixes."""
    match = re.fullmatch(
        r"([0-9]+(?:\.[0-9]+)*)(?:(?:a|b|rc)[0-9]+)?(?:\.post[0-9]+)?(?:\.?dev[0-9]+)?"
        r"(?:\+[a-zA-Z0-9]+(?:[.-][a-zA-Z0-9]+)*)?", version,
    )
    if not match:
        raise ValueError(f"Cannot parse Windows version: {version!r}")
    parts = [int(part) for part in match[1].split(".")[:4]]
    if any(part > 65535 for part in parts):
        raise ValueError(f"Windows version components must fit in 16 bits: {version!r}")
    return tuple(parts + [0] * (4 - len(parts)))
