"""Word wildcard counts {n,m} follow the list separator of Word's locale (#56)."""

import pytest
import pywintypes

from office_live import wd_common


@pytest.mark.parametrize("pattern, sep, expected", [
    ("В[0-9]{1,2}", ";", "В[0-9]{1;2}"),
    ("В[0-9]{1,2}", ",", "В[0-9]{1,2}"),
    ("[0-9]{1;2}", ",", "[0-9]{1,2}"),
    ("a{2,}b{3}", ";", "a{2;}b{3}"),
    ("[{1,2}]x{1,2}", ";", "[{1,2}]x{1;2}"),
    ("\\{1,2\\}", ";", "\\{1,2\\}"),
    ("{a,b}", ";", "{a,b}"),
    ("{1,2", ";", "{1,2"),
])
def test_localize_wildcards(pattern, sep, expected):
    assert wd_common.localize_wildcards(pattern, sep) == expected


def test_bad_wildcard_error_names_separator():
    exc = pywintypes.com_error(-2147352567, "Exception occurred.",
                               (0, "Microsoft Word", "Поле 'Найти' содержит неверный шаблон поиска", None, 0, -2146822728), None)
    err = wd_common.wildcard_error(exc, "В[0-9]{1;2", ";")
    assert err is not None and "[0-9]{1;2}" in str(err) and "0x800A15B8" in str(err)


def test_other_com_errors_are_not_rewritten():
    exc = pywintypes.com_error(-2147352567, "Exception occurred.", (0, "Word", "Busy", None, 0, -2147418111), None)
    assert wd_common.wildcard_error(exc, "x", ";") is None
