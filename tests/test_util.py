"""Чистая логика без Office: адреса, цвета, значения, форматы чисел, разбор текста."""

import datetime

import pytest

from office_live import util
from office_live.errors import ToolError


def test_col_letters_roundtrip():
    assert util.col_letter(1) == "A"
    assert util.col_letter(26) == "Z"
    assert util.col_letter(27) == "AA"
    assert util.col_letter(16384) == "XFD"
    for n in (1, 26, 27, 52, 702, 703, 16384):
        assert util.col_number(util.col_letter(n)) == n
    with pytest.raises(ValueError):
        util.col_letter(0)


@pytest.mark.parametrize(
    "addr,expected",
    [
        ("B2", (2, 2, 2, 2)),
        ("$B$2:$D$5", (2, 2, 5, 4)),
        ("D5:B2", (2, 2, 5, 4)),
        ("A:C", (1, 1, util.MAX_ROWS, 3)),
        ("3:7", (3, 1, 7, util.MAX_COLS)),
    ],
)
def test_parse_a1(addr, expected):
    assert util.parse_a1(addr) == expected


def test_parse_a1_rejects_garbage():
    with pytest.raises(ValueError):
        util.parse_a1("not a range")


def test_split_sheet_ref_and_quote():
    assert util.split_sheet_ref("A1:B2") == (None, "A1:B2")
    assert util.split_sheet_ref("Data!A1") == ("Data", "A1")
    assert util.split_sheet_ref("'My Sheet'!A1:B2") == ("My Sheet", "A1:B2")
    assert util.split_sheet_ref("'It''s'!C3") == ("It's", "C3")
    assert util.quote_sheet("Data") == "Data"
    assert util.quote_sheet("My Sheet") == "'My Sheet'"
    assert util.quote_sheet("It's") == "'It''s'"


def test_a1_range_collapses_single_cell():
    assert util.a1_range(2, 2, 2, 2) == "B2"
    assert util.a1_range(1, 1, 3, 4) == "A1:D3"


def test_parse_row_spec():
    assert util.parse_row_spec("3-5,8") == [3, 4, 5, 8]
    assert util.parse_row_spec("5-3") == [3, 4, 5]
    with pytest.raises(ToolError):
        util.parse_row_spec("1-100", limit=10)


def test_colors():
    assert util.parse_color("#FF0000") == 0x0000FF  # COLORREF = BGR
    assert util.parse_color("white") == 0xFFFFFF
    assert util.parse_color("WHITE") == 0xFFFFFF
    assert util.parse_color("#abc") == util.parse_color("#AABBCC")
    assert util.color_to_hex(util.parse_color("#1F4E78")) == "#1F4E78"
    assert util.color_to_hex(-1) is None and util.color_to_hex(None) is None
    with pytest.raises(ToolError):
        util.parse_color("not-a-color")


def test_excel_error_codes_become_text():
    assert util.norm_value(-2146826281) == "#DIV/0!"
    assert util.norm_value(-2146826246) == "#N/A"
    assert util.norm_value(-2146826265) == "#REF!"
    assert util.norm_value(42) == 42  # обычные числа не трогаем


def test_norm_value_types():
    assert util.norm_value(5.0) == 5 and isinstance(util.norm_value(5.0), int)
    assert util.norm_value(2.5) == 2.5
    assert util.norm_value(float("nan")) is None
    assert util.norm_value(datetime.datetime(2026, 10, 4)) == "2026-10-04"
    assert util.norm_value(datetime.datetime(2026, 10, 4, 13, 5, 9)) == "2026-10-04 13:05:09"
    assert util.norm_value(datetime.datetime(1899, 12, 30, 8, 30)) == "08:30:00"  # «чистое» время Excel
    assert util.norm_value(None) is None and util.norm_value(True) is True


def test_grid_shapes():
    assert util.to_grid(5) == [[5]]
    assert util.to_grid((1, 2)) == [[1, 2]]
    assert util.to_grid(((1, 2), (3, 4))) == [[1, 2], [3, 4]]


def test_to_com_grid_pads_and_wraps():
    assert util.to_com_grid([1, 2, 3]) == ((1, 2, 3),)
    assert util.to_com_grid([[1], [2, 3]]) == ((1, None), (2, 3))
    assert util.to_com_grid("x") == (("x",),)
    with pytest.raises(ToolError):
        util.to_com_grid([])


@pytest.mark.parametrize(
    "text,decimal,expected",
    [
        ("1 234,50", ",", 1234.5),
        ("1 234,50", ",", 1234.5),
        ("1,234.50", ".", 1234.5),
        ("1.234,50", ",", 1234.5),
        ("12%", ",", 0.12),
        ("8,5%", ",", 0.085),
        ("-12", ",", -12),
        ("(12)", ",", -12),
        ("12", ",", 12),
        ("1,234", ".", 1234),       # при десятичной точке это тысячи
        ("1,234", ",", 1.234),      # при десятичной запятой — дробь
        ("0,5", ",", 0.5),
    ],
)
def test_smart_number(text, decimal, expected):
    assert util.smart_number(text, decimal) == pytest.approx(expected)


@pytest.mark.parametrize("text", ["007", "0012", "", "abc", "12abc", "1-2", "+7 999 123", "12345678901234567890", None])
def test_smart_number_leaves_codes_as_text(text):
    assert util.smart_number(text, ",") is None


@pytest.mark.parametrize(
    "text",
    [
        "12.03.2026",      # дата, а не 12032026
        "1.2.3", "1.2.3.4",  # номера разделов
        "1,23,4", "12,34,56",  # неправильные группы разрядов
        "12 34", "1 23 456",   # пробельные группы не по 3 цифры
        "1,234.5,6", "1.234,56,7", ",5", "5,", "1..2", "1 234 ,5",
        "0,123,456",       # группа не может начинаться с нуля
        "-", "--5", "(", "()", "%", ".",
    ],
)
def test_smart_number_rejects_ambiguous_or_malformed_groupings(text):
    assert util.smart_number(text, ",") is None
    assert util.smart_number(text, ".") is None


@pytest.mark.parametrize(
    "text,decimal,expected",
    [
        ("1 234 567", ",", 1234567), ("12 345,6", ",", 12345.6), ("1,234,567", ".", 1234567), ("1.234.567", ",", 1234567),
        ("1,234,567.89", ".", 1234567.89), ("1.234.567,89", ",", 1234567.89), ("-1 234", ",", -1234), ("+5", ",", 5),
        ("0.5", ".", 0.5), ("0,123", ",", 0.123), ("0,123", ".", 0.123), ("100%", ",", 1.0), ("(1 000)", ",", -1000),
    ],
)
def test_smart_number_accepts_well_formed_groupings(text, decimal, expected):
    assert util.smart_number(text, decimal) == pytest.approx(expected)


def test_clean_word_text():
    assert util.clean_word_text("cell\r\x07") == "cell"
    assert util.clean_word_text("a\rb") == "a\nb"
    assert util.clean_word_text("line\x0bbreak") == "line\nbreak"
    assert util.clean_word_text(None) == ""
    assert util.to_word_text("a\nb\r\nc") == "a\rb\rc"


@pytest.mark.parametrize(
    "fmt",
    ["#,##0.00", "0.0%", "dd.mm.yyyy", '0.0 "kg"', "yyyy-mm-dd hh:mm", "0.00E+00", "$#,##0.00", "#,##0", "0", "@", "General", '#,##0.00 "₽"'],
)
def test_number_format_roundtrip_ru(fmt):
    """Региональные настройки RU: десятичная ',', тысячи — неразрывный пробел."""
    local = util.localize_number_format(fmt, ",", " ")
    assert util.delocalize_number_format(local, ",", " ") == fmt


def test_number_format_localization_details():
    assert util.localize_number_format("0.00", ",", " ") == "0,00"
    assert util.localize_number_format("#,##0.00", ",", " ") == "# ##0,00"
    assert util.localize_number_format("dd.mm.yyyy", ",", " ") == "dd.mm.yyyy"  # точки в датах — литералы
    assert util.localize_number_format("0.0 \"kg\"", ",", " ") == '0,0\\ "kg"'  # пробел после цифр иначе масштабирует
    assert util.localize_number_format("#,##0.00", ".", ",") == "#,##0.00"  # англ. настройки — без изменений


class _Ole:
    """COM-интерфейс, который понимает только английские коды с LCID 1033 (как русский Excel, проверено вживую)."""

    def __init__(self, owner, reject=False):
        self.owner, self.reject, self.calls = owner, reject, []

    def GetIDsOfNames(self, name):
        assert name == "NumberFormat"
        return 193

    def Invoke(self, dispid, lcid, flags, result, *args):
        import pythoncom
        import pywintypes

        self.calls.append((dispid, lcid, flags, *args))
        if self.reject:
            raise pywintypes.com_error(-2146827284, "Exception occurred.", None, None)
        if flags == pythoncom.DISPATCH_PROPERTYPUT:
            self.owner.english = args[0]
            return None
        return self.owner.english


class _Target:
    def __init__(self, reject=False):
        self.english, self.NumberFormat = "General", "Основной"
        self._oleobj_ = _Ole(self, reject)


_RU = type("App", (), {"International": (None, None, ",", " ", ";")})()


@pytest.mark.parametrize("fmt,expected", [("dd.mm.yyyy", "dd.mm.yyyy"), ("date", "dd.mm.yyyy"), ("general", "General"),
                                          ("[Red]#,##0.00", "[Red]#,##0.00"), ("[h]:mm", "[h]:mm")])
def test_number_format_is_written_in_english_with_lcid_1033(fmt, expected):
    # Демо 2026-10-05: русский Excel записывал 'dd.mm.yyyy' как текст (ячейка показывала 'dd.mm.yyyy' или ####).
    import pythoncom

    from office_live.excel_format import NUMBER_FORMATS
    from office_live.xl_common import read_number_format, set_number_format

    target = _Target()
    set_number_format(_RU, target, fmt, NUMBER_FORMATS)
    assert target._oleobj_.calls == [(193, 1033, pythoncom.DISPATCH_PROPERTYPUT, expected)]
    assert target.NumberFormat == "Основной"  # локальный путь не использовался
    assert read_number_format(_RU, target) == expected


def test_local_number_formats_and_fakes_keep_the_local_path():
    from office_live.xl_common import set_number_format

    target = _Target()
    set_number_format(_RU, target, "ДД.ММ.ГГГГ")  # уже по-местному: английский разбор сделал бы буквы текстом
    assert target.NumberFormat == "ДД.ММ.ГГГГ" and target._oleobj_.calls == []
    set_number_format(_RU, target, '0.0 "кг"')  # кириллица в кавычках — литерал, код английский
    assert target.english == '0.0 "кг"'
    rejected = _Target(reject=True)
    set_number_format(_RU, rejected, "0.00")  # отказ LCID 1033 -> прежняя локальная запись
    assert rejected.NumberFormat == "0,00"
    fake = type("Fake", (), {"NumberFormat": "General"})()
    set_number_format(_RU, fake, "#,##0.00")
    assert fake.NumberFormat == "# ##0,00"


def test_cm_to_points():
    assert util.cm_to_points(2.54) == pytest.approx(72.0)


def test_truncate():
    assert util.truncate("abcdef", 3) == ("abc", True)
    assert util.truncate("abc", 3) == ("abc", False)
    assert util.truncate("abc", 0) == ("abc", False)
