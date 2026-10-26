"""word_find/find_all stay bounded when Word reports hits outside the searched range (#55)."""

from office_live import wd_common, word_core


class FakeRange:
    def __init__(self, text, start, end, stuck_hit=None, story_type=1):
        self.text, self.Start, self.End = text, start, end
        self.stuck_hit = stuck_hit  # Word's header quirk: the same hit, whatever the range
        self.StoryType = story_type
        self.executes = 0

    @property
    def Duplicate(self):
        return FakeRange(self.text, self.Start, self.End, self.stuck_hit, self.StoryType)

    @property
    def Find(self):
        return self

    @property
    def Text(self):
        return self.text[self.Start:self.End]

    def ClearFormatting(self):
        pass

    def Execute(self, text, *args):
        self.executes += 1
        if self.stuck_hit is not None:
            self.Start, self.End = self.stuck_hit
            return True
        at = self.text.find(text, self.Start, self.End)
        if at < 0:
            return False
        self.Start, self.End = at, at + len(text)
        return True

    def SetRange(self, start, end):
        self.Start, self.End = start, end


def test_find_all_returns_every_hit_in_order():
    story = FakeRange("a 1660 b 1660 c", 0, 15)
    assert [(r.Start, r.End) for r in wd_common.find_all(story, "1660")] == [(2, 6), (9, 13)]


def test_find_all_stops_when_word_repeats_a_hit_before_the_range():
    # Before the fix the hit at 2..6 came back after SetRange(6, 30) and the loop ran until limit=100000.
    story = FakeRange("x" * 30, 0, 30, stuck_hit=(2, 6))
    assert [(r.Start, r.End) for r in wd_common.find_all(story, "1660")] == [(2, 6)]


def test_find_all_ignores_a_hit_outside_the_story():
    story = FakeRange("x" * 30, 10, 20, stuck_hit=(25, 29))
    assert list(wd_common.find_all(story, "1660")) == []


def test_word_find_returns_partial_result_after_time_budget(monkeypatch):
    doc = type("Doc", (), {"Name": "doc.docx"})()
    stories = [FakeRange("1660 1660", 0, 9, story_type=7), FakeRange("1660", 0, 4, story_type=9)]
    clock = iter([0.0, 1.0, 2.0, 999.0, 999.0, 999.0])
    monkeypatch.setattr(word_core, "pick_document", lambda name: (None, doc))
    monkeypatch.setattr(word_core, "story_ranges", lambda d, scope: stories)
    monkeypatch.setattr(word_core.time, "monotonic", lambda: next(clock))
    result = word_core.word_find("1660", "doc.docx", scope="all")
    assert result["partial"] is True
    assert (result["total_matches"], result["stories_searched"], result["stories_total"]) == (1, 0, 2)
