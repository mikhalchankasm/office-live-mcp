"""Настройки сервера из переменных окружения OFFICE_LIVE_*.

OFFICE_LIVE_MODE            full (по умолчанию) | readonly — в readonly пишущие инструменты
                            вообще не регистрируются, книги/документы открываются только на чтение
OFFICE_LIVE_TOOLSETS        all (по умолчанию) | core | excel | word | список групп через запятую:
                            excel_core, excel_format, excel_analysis, word_core, word_tables,
                            word_layout, bridge  (меньше групп = меньше токенов в контексте агента)
OFFICE_LIVE_ALLOWED_DIRS    каталоги через ';' — открывать/сохранять/вставлять файлы только отсюда
OFFICE_LIVE_ALLOW_EVAL      1 — зарегистрировать office_run_python (произвольный Python с доступом к COM)
OFFICE_LIVE_AUDIT_LOG       путь к jsonl-журналу всех пишущих вызовов
OFFICE_LIVE_STRICT_TARGET   1 — пишущие инструменты требуют явное имя книги/документа (без «активного»)
OFFICE_LIVE_BUSY_TIMEOUT    секунды ожидания, пока Office «занят» (редактирование ячейки, диалог); по умолчанию 20
"""

import os
import sys
from dataclasses import dataclass
from pathlib import Path

GROUPS = (
    "excel_core",
    "excel_format",
    "excel_analysis",
    "word_core",
    "word_tables",
    "word_layout",
    "bridge",
)

PRESETS = {
    "all": GROUPS,
    "core": ("excel_core", "word_core", "bridge"),
    "excel": ("excel_core", "excel_format", "excel_analysis"),
    "word": ("word_core", "word_tables", "word_layout"),
}

# виды инструментов: read/ui не меняют содержимое; open открывает файл; остальные меняют документы или диск
READ_KINDS = frozenset({"read", "ui", "open"})
ALL_KINDS = READ_KINDS | {"write", "destructive", "save"}


@dataclass(frozen=True)
class Settings:
    mode: str = "full"
    groups: frozenset = frozenset(GROUPS)
    allowed_dirs: tuple = ()
    allow_eval: bool = False
    audit_log: Path | None = None
    strict_target: bool = False
    busy_timeout: float = 20.0

    @property
    def readonly(self) -> bool:
        return self.mode == "readonly"

    def enabled(self, group: str, kind: str) -> bool:
        if group == "eval":
            return self.allow_eval and not self.readonly
        if group not in self.groups:
            return False
        if self.readonly and kind not in READ_KINDS:
            return False
        return True


def _flag(value: str | None) -> bool:
    return (value or "").strip().lower() in {"1", "true", "yes", "on"}


def load(env=None) -> Settings:
    env = os.environ if env is None else env

    mode = (env.get("OFFICE_LIVE_MODE") or "full").strip().lower()
    if mode not in {"full", "readonly"}:
        print(f"[office-live] неизвестный OFFICE_LIVE_MODE={mode!r}, использую 'readonly' из осторожности", file=sys.stderr)
        mode = "readonly"

    raw = (env.get("OFFICE_LIVE_TOOLSETS") or "all").strip().lower()
    groups: set[str] = set()
    for token in (t.strip() for t in raw.replace(";", ",").split(",")):
        if not token:
            continue
        if token in PRESETS:
            groups.update(PRESETS[token])
        elif token in GROUPS:
            groups.add(token)
        else:
            print(f"[office-live] неизвестный набор инструментов {token!r} пропущен", file=sys.stderr)
    if not groups:
        groups = set(GROUPS)

    dirs = []
    for part in (env.get("OFFICE_LIVE_ALLOWED_DIRS") or "").split(";"):
        part = part.strip().strip('"')
        if part:
            dirs.append(Path(os.path.abspath(os.path.expandvars(part))))

    audit = (env.get("OFFICE_LIVE_AUDIT_LOG") or "").strip().strip('"')
    try:
        timeout = float(env.get("OFFICE_LIVE_BUSY_TIMEOUT") or 20)
    except ValueError:
        timeout = 20.0

    return Settings(
        mode=mode,
        groups=frozenset(groups),
        allowed_dirs=tuple(dirs),
        allow_eval=_flag(env.get("OFFICE_LIVE_ALLOW_EVAL")),
        audit_log=Path(audit) if audit else None,
        strict_target=_flag(env.get("OFFICE_LIVE_STRICT_TARGET")),
        busy_timeout=max(0.0, timeout),
    )


SETTINGS = load()
