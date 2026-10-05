"""Общие фикстуры и помощники живых тестов (реальные Excel/Word через настоящий MCP-протокол).

Фикстуры создают СВОИ черновые книги/документы и закрывают только их (discard); приложения не закрываются, кроме Word,
если его запустили сами тесты и в нём не осталось документов.
"""

import tempfile
import os
import uuid

import pytest

from tests.live.mcpclient import Stdio, ToolFailed
from tests.live.office_cleanup import quit_word_if_idle, word_running


@pytest.fixture(scope="session")
def srv():
    word_was_running = word_running()  # Word пользователя не закрываем, даже пустой
    c = Stdio()
    yield c
    c.close()
    quit_word_if_idle(started_by_tests=not word_was_running)


# уникальна для этого запуска: файлы других запусков и пользователя не совпадут. Через окружение — потому что модуль
# грузится дважды (pytest как conftest и тесты как tests.live.conftest), а метка должна быть одна.
RUN_TAG = os.environ.setdefault("OFFICE_LIVE_TEST_RUN_TAG", f"ol_pytest_{uuid.uuid4().hex[:8]}")


def _in_run_dir(path) -> bool:
    """Файл лежит во временной папке ЭТОГО запуска (туда тесты делают Save As): путь уникален, чужим быть не может."""
    return bool(path) and RUN_TAG in path


def _full(item) -> str:
    return item["path"] + "\\" + item["name"] if item["path"] else item["name"]


@pytest.fixture
def wb(srv):
    name = srv.call("excel_new_workbook", sheets=["Data"])["workbook"]
    marker = f"{RUN_TAG}_{uuid.uuid4().hex[:8]}"
    # имя несохранённой книги («Книга3») может совпасть с книгой пользователя в другом экземпляре Excel:
    # свою узнаём по уникальной метке, а не по имени
    srv.call("excel_manage_names", workbook=name, action="add", name=marker, formula="=1")
    yield name
    close_own_workbooks(srv, name, marker)


def close_own_workbooks(srv, name, marker):
    """Закрывает (discard) только книги этого теста: файлы в папке запуска и несохранённую книгу с нашей меткой."""
    for w in srv.call("excel_list_workbooks")["workbooks"]:
        if _in_run_dir(w["path"]):
            srv.call("excel_close_workbook", workbook=_full(w), discard=True)
        elif not w["path"] and w["name"] == name and _excel_marked(srv, name, marker):
            srv.call("excel_close_workbook", workbook=name, discard=True)


def _excel_marked(srv, name, marker) -> bool:
    try:  # имя в нескольких экземплярах -> сервер откажет; тогда не закрываем ничего
        names = srv.call("excel_manage_names", workbook=name, action="list")["names"]
    except ToolFailed:
        return False
    return any(n["name"] == marker for n in names)


@pytest.fixture
def doc(srv):
    name = srv.call("word_new_document")["document"]
    marker = f"{RUN_TAG}_{uuid.uuid4().hex[:8]}"
    srv.call("word_document_properties", document=name, action="set", properties={"Comments": marker})
    yield name
    close_own_documents(srv, name, marker)


def close_own_documents(srv, name, marker):
    for d in srv.call("word_list_documents")["documents"]:
        if _in_run_dir(d["path"]):
            srv.call("word_close_document", document=_full(d), discard=True)
        elif not d["path"] and d["name"] == name and _word_marked(srv, name, marker):
            srv.call("word_close_document", document=name, discard=True)


def _word_marked(srv, name, marker) -> bool:
    try:
        props = srv.call("word_document_properties", document=name)["properties"]
    except ToolFailed:
        return False
    return props.get("Comments") == marker


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


