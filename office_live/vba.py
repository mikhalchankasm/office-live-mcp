"""Чтение исходного кода макросов из vbaProject.bin без запуска Office (MS-OVBA: контейнер сжатия + поток dir).

Нужен, чтобы агент мог «разобрать» файл .xlsm/.docm: увидеть, какие процедуры есть в модулях, не открывая файл
и не требуя включённого «Доверять доступ к объектной модели VBA». Код только читается, не исполняется.
"""

import io
import re
import struct

try:
    import olefile
except ImportError:  # необязательная зависимость: без неё покажем только имена модулей
    olefile = None

_CODEPAGES = {1251: "cp1251", 1252: "cp1252", 1250: "cp1250", 65001: "utf-8"}


def decompress(data: bytes) -> bytes:
    """Алгоритм сжатия MS-OVBA 2.4.1."""
    if not data or data[0] != 1:
        raise ValueError("not an MS-OVBA compressed container")
    out = bytearray()
    pos = 1
    while pos < len(data):
        header = struct.unpack_from("<H", data, pos)[0]
        chunk_end = min(pos + (header & 0x0FFF) + 3, len(data))
        compressed = bool(header & 0x8000)
        pos += 2
        if not compressed:
            out += data[pos:pos + 4096]
            pos += 4096
            continue
        chunk_start = len(out)
        while pos < chunk_end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= chunk_end:
                    break
                if not (flags >> bit) & 1:
                    out.append(data[pos])
                    pos += 1
                else:
                    token = struct.unpack_from("<H", data, pos)[0]
                    pos += 2
                    diff = len(out) - chunk_start
                    bits = max(4, (diff - 1).bit_length()) if diff > 1 else 4
                    bits = min(bits, 12)
                    length_mask = 0xFFFF >> bits
                    length = (token & length_mask) + 3
                    offset = ((token & ~length_mask & 0xFFFF) >> (16 - bits)) + 1
                    for _ in range(length):
                        out.append(out[-offset])
        pos = chunk_end
    return bytes(out)


def _parse_dir(raw: bytes):
    """Поток dir -> (кодовая страница, [модули])."""
    d = decompress(raw)
    pos, codepage, modules, cur = 0, 1252, [], None
    while pos + 6 <= len(d):
        rec_id, size = struct.unpack_from("<HI", d, pos)
        pos += 6
        if rec_id == 0x0009:  # PROJECTVERSION: фиксированная длина 6 после поля размера
            size = 6
        data = d[pos:pos + size]
        pos += size
        if rec_id == 0x0003 and size == 2:
            codepage = struct.unpack("<H", data)[0]
        elif rec_id == 0x0019:  # MODULENAME
            cur = {"name": data.decode(_CODEPAGES.get(codepage, "cp1252"), "replace"), "stream": None, "offset": 0, "kind": "standard"}
            modules.append(cur)
        elif cur is not None and rec_id == 0x001A:  # MODULESTREAMNAME
            cur["stream"] = data.decode(_CODEPAGES.get(codepage, "cp1252"), "replace")
        elif cur is not None and rec_id == 0x0031 and size == 4:  # MODULEOFFSET
            cur["offset"] = struct.unpack("<I", data)[0]
        elif cur is not None and rec_id == 0x0021:
            cur["kind"] = "standard"
        elif cur is not None and rec_id == 0x0022:
            cur["kind"] = "class_or_document"
    return codepage, modules


_PROC = re.compile(r"^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)", re.IGNORECASE | re.MULTILINE)


def read_vba_project(project_bin: bytes, include_source: bool = False, module_filter: str = "", max_chars: int = 20000) -> dict:
    """Описание VBA-проекта: модули, процедуры, (по запросу) исходный текст."""
    info: dict = {"modules": [], "readable": False}
    names_hint = sorted({m.decode("latin-1") for m in re.findall(rb"(?:Module|Class|Document)=([A-Za-z0-9_]+)", project_bin)})
    if olefile is None:
        info["note"] = "Install 'olefile' (pip install olefile) to read macro source."
        info["module_names"] = names_hint
        return info
    try:
        ole = olefile.OleFileIO(io.BytesIO(project_bin))
    except Exception as exc:  # noqa: BLE001
        info["note"] = f"Could not parse the VBA container: {exc}"
        info["module_names"] = names_hint
        return info
    try:
        codepage, modules = _parse_dir(ole.openstream("VBA/dir").read())
        enc = _CODEPAGES.get(codepage, "cp1252")
        for m in modules:
            entry = {"name": m["name"], "kind": m["kind"]}
            try:
                stream = ole.openstream("VBA/" + (m["stream"] or m["name"])).read()
                src = decompress(stream[m["offset"]:]).decode(enc, "replace").replace("\r\n", "\n")
            except Exception as exc:  # noqa: BLE001
                entry["error"] = f"could not read module: {exc}"
                info["modules"].append(entry)
                continue
            body = "\n".join(line for line in src.split("\n") if not line.startswith("Attribute VB_"))
            entry["lines"] = body.count("\n") + 1 if body.strip() else 0
            entry["procedures"] = [f"{k.strip()} {n}" for k, n in _PROC.findall(body)]
            if include_source and (not module_filter or module_filter.lower() == m["name"].lower()):
                entry["source"] = body[:max_chars]
                entry["source_truncated"] = len(body) > max_chars
            info["modules"].append(entry)
        info["readable"] = True
        info["codepage"] = codepage
    except Exception as exc:  # noqa: BLE001
        info["note"] = f"Could not read the VBA project: {exc}"
        info["module_names"] = names_hint
    finally:
        ole.close()
    return info
