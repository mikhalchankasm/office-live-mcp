"""Общие фикстуры и помощники живых тестов (реальные Excel/Word через настоящий MCP-протокол).

Фикстуры создают СВОИ черновые книги/документы и закрывают только их (discard); приложения не закрываются, кроме Word,
если его запустили сами тесты и в нём не осталось документов.
"""

import tempfile
import uuid

import pytest

from tests.live.mcpclient import Stdio
from tests.live.office_cleanup import quit_word_if_idle, word_running


@pytest.fixture(scope="session")
def srv():
    word_was_running = word_running()  # Word пользователя не закрываем, даже пустой
    c = Stdio()
    yield c
    c.close()
    quit_word_if_idle(started_by_tests=not word_was_running)


RUN_TAG = f"ol_pytest_{uuid.uuid4().hex[:8]}"  # уникальна для этого запуска: файлы других запусков и пользователя не совпадут


def _is_ours(name, original, path):
    """Наш файл: исходное имя черновика или файл в временной папке ЭТОГО запуска (имя меняется после Save As)."""
    return name == original or (bool(path) and RUN_TAG in path)


@pytest.fixture
def wb(srv):
    name = srv.call("excel_new_workbook", sheets=["Data"])["workbook"]
    yield name
    for w in srv.call("excel_list_workbooks")["workbooks"]:
        if _is_ours(w["name"], name, w["path"]):
            srv.call("excel_close_workbook", workbook=w["path"] + "\\" + w["name"] if w["path"] else w["name"], discard=True)


@pytest.fixture
def doc(srv):
    name = srv.call("word_new_document")["document"]
    yield name
    for d in srv.call("word_list_documents")["documents"]:
        if _is_ours(d["name"], name, d["path"]):
            srv.call("word_close_document", document=d["path"] + "\\" + d["name"] if d["path"] else d["name"], discard=True)


@pytest.fixture
def tmp():
    d = tempfile.mkdtemp(prefix=RUN_TAG + "_")
    yield d
    import shutil

    shutil.rmtree(d, ignore_errors=True)


def assert_real_picture(png: bytes, min_size=(60, 20)):
    """PNG декодируется, имеет разумный размер и не залит одним цветом (пустой снимок вне видимой области)."""
    import io

    from PIL import Image as PILImage

    assert png[:8] == b"\x89PNG\r\n\x1a\n"
    im = PILImage.open(io.BytesIO(png)).convert("L")
    assert im.width >= min_size[0] and im.height >= min_size[1], im.size
    hist = im.histogram()
    assert max(hist) / float(im.width * im.height) < 0.97, "the picture is a single flat colour"


def read(srv, wb, cells, sheet="Data", **kw):
    return srv.call("excel_read_range", workbook=wb, sheet=sheet, cells=cells, **kw)["values"]


