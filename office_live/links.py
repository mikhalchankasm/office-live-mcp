"""Pure, bounded officelive URI codec. No COM, file opening or command execution."""

import ntpath
import re
import unicodedata
from urllib.parse import quote, unquote_to_bytes

from .errors import ToolError
from .util import MAX_COLS, MAX_ROWS, parse_a1

MAX_URI = 2048
MAX_INDEX = 1_000_000
MAX_LINKS = 20
FIELDS = {"excel": {"book", "sheet", "range"}, "word": {"doc", "paragraph", "paragraphs", "table", "bookmark", "span"}}


def _fail():
    raise ToolError("Invalid Office Live link. Only a local, already open Excel/Word target is allowed.")


def _text(value, limit):
    if (not isinstance(value, str) or not value or len(value) > limit or value != value.strip()
            or any(unicodedata.category(c).startswith("C") for c in value)):
        _fail()


def file_target(value):
    _text(value, 1024)
    # No UNC/device/URL/relative paths, traversal, ADS, quotes or shell placeholders.
    value = value.replace("/", "\\")
    drive, tail = ntpath.splitdrive(value)
    if drive:
        if not re.fullmatch(r"[A-Za-z]:", drive) or not tail.startswith("\\"):
            _fail()
        tail = tail[1:]
    elif "\\" in tail or ":" in tail:
        _fail()
    parts = tail.split("\\")
    if any(not p or p in {".", ".."} or p.endswith((".", " ")) or any(c in p for c in '<>:"|?*%') for p in parts):
        _fail()
    return value


def a1(value):
    _text(value, 64)
    cell = r"\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}"
    col = r"\$?[A-Za-z]{1,3}"
    row = r"\$?[1-9][0-9]{0,6}"
    if not re.fullmatch(rf"(?:{cell}(?::{cell})?|{col}:{col}|{row}:{row})", value):
        _fail()
    r1, c1, r2, c2 = parse_a1(value)
    if not (1 <= r1 <= r2 <= MAX_ROWS and 1 <= c1 <= c2 <= MAX_COLS):
        _fail()
    # parse_a1 sorts reversed bounds; links must not silently reinterpret them.
    if ":" in value:
        left, right = value.split(":")
        lr, lc, _, _ = parse_a1(left)
        rr, rc, _, _ = parse_a1(right)
        if lr > rr or lc > rc:
            _fail()
    return value.replace("$", "").upper()


def _index(value):
    if not re.fullmatch(r"[1-9][0-9]{0,6}", value) or int(value) > MAX_INDEX:
        _fail()
    return int(value)


def validate(app, params):
    if app not in FIELDS or set(params) - FIELDS[app]:
        _fail()
    for value in params.values():
        _text(value, 1024)
    file_target(params.get("book" if app == "excel" else "doc", ""))
    if app == "excel":
        if "sheet" in params:
            sheet = params["sheet"]
            if len(sheet) > 31 or any(c in sheet for c in '[]:*?/\\') or sheet.startswith("'") or sheet.endswith("'"):
                _fail()
        if "range" in params:
            if "sheet" not in params:  # stable target; never depend on the currently active sheet
                _fail()
            a1(params["range"])
    else:
        selectors = set(params) - {"doc"}
        if len(selectors) > 1:
            _fail()
        for key in selectors:
            value = params[key]
            if key in {"paragraph", "table"}:
                _index(value)
            elif key == "span":
                if not re.fullmatch(r"(?:0|[1-9][0-9]{0,9})-(?:0|[1-9][0-9]{0,9})", value, flags=re.ASCII):
                    _fail()
                start, end = map(int, value.split("-"))
                if not 0 <= start <= end <= 2147483647:
                    _fail()
            elif key == "paragraphs":
                parts = value.split("-")
                if len(parts) != 2 or _index(parts[0]) > _index(parts[1]):
                    _fail()
            elif len(value) > 40 or not re.fullmatch(r"[^\W\d][\w]*", value):
                _fail()
    return app, params


def parse(uri):
    if not isinstance(uri, str) or len(uri) > MAX_URI:
        _fail()
    match = re.fullmatch(r"officelive://(excel|word)/open\?([^#]+)", uri, flags=re.ASCII)
    if not match:
        _fail()
    app, query = match.groups()
    params = {}
    for field in query.split("&"):
        key, sep, value = field.partition("=")
        if not sep or key not in FIELDS[app] or key in params or not re.fullmatch(r"(?:[A-Za-z0-9_.~-]|%[0-9A-Fa-f]{2})+", value):
            _fail()
        try:
            params[key] = unquote_to_bytes(value).decode("utf-8", "strict")
        except UnicodeError:
            _fail()
    return validate(app, params)


def build(app, **params):
    params = {k: str(v) for k, v in params.items() if v is not None and v != ""}
    validate(app, params)
    uri = f"officelive://{app}/open?" + "&".join(f"{k}={quote(v, safe='')}" for k, v in params.items())
    if len(uri) > MAX_URI:
        _fail()
    return uri


def markdown(link):
    label = re.sub(r"([\\`*_{}\[\]()<>!|])", r"\\\1", link["label"])
    return f"[{label}]({link['uri']})"
