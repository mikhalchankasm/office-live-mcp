"""A transient GetIDsOfNames failure is retried instead of surfacing as '<unknown>.Name' (#57)."""

import pytest

from office_live import com


class Ole:
    def __init__(self, known=True):
        self.known = known

    def GetIDsOfNames(self, lcid, name):
        if not self.known:
            raise AttributeError(name)
        return 1


class Flaky:
    """getattr fails once, as CDispatch does when Word rejects the first name lookup."""

    def __init__(self, failures=1, known=True):
        object.__setattr__(self, "_oleobj_", Ole(known))
        object.__setattr__(self, "failures", failures)

    def __getattr__(self, name):
        if self.failures:
            object.__setattr__(self, "failures", self.failures - 1)
            raise AttributeError(f"<unknown>.{name}")
        return lambda *args: f"{name}{args}"


def test_transient_lookup_failure_is_retried():
    assert com.Proxy(Flaky()).StartCustomRecord("x") == "StartCustomRecord('x',)"


def test_unknown_name_still_raises():
    with pytest.raises(AttributeError, match="StartCustomRecord"):
        com.Proxy(Flaky(failures=1, known=False)).StartCustomRecord("x")
