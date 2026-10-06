"""In-memory Word fragment, revision and comment contracts. Never creates COM objects."""

from datetime import datetime
from types import SimpleNamespace as NS

from tests.history_fakes import Collection
from tests.office_operations_fakes import WordRange, error, word


def units(text):
    return len(text.encode("utf-16-le")) // 2


def font(**changes):
    values = {"Name": "Calibri", "Size": 11, "Bold": 0, "Italic": 0, "Underline": 0, "Color": 0, **changes}
    obj = NS(**values)
    obj.Duplicate = NS(**values)
    return obj


class EditRange(WordRange):
    def __init__(self, doc, start, end):
        super().__init__(doc, start, end)
        self.StoryType = doc.story
        for member in ("Footnotes", "Endnotes", "Fields", "Hyperlinks", "ContentControls", "InlineShapes", "Revisions"):
            objects = getattr(doc, "range_objects", {}).get(member, [])
            setattr(self, member, Collection(objects))
        self.Style, self.ListFormat = doc.style, NS(ListString=doc.list_string)
        self.Font = font(Bold=9999999 if doc.mixed and end - start > 1 else 0)
        self.HighlightColorIndex = 0

    def __deepcopy__(self, memo):
        return EditRange(self.doc, self.Start, self.End)

    @property
    def Text(self):
        raw = self.doc.Content.Text.encode("utf-16-le")
        return raw[self.Start * 2:self.End * 2].decode("utf-16-le")

    @Text.setter
    def Text(self, value):
        doc = self.doc
        doc.calls.append(("edit", self.Start, self.End, value, doc.TrackRevisions))
        raw = doc.Content.Text.encode("utf-16-le")
        doc.Content.Text = (raw[:self.Start * 2] + value.encode("utf-16-le") + raw[self.End * 2:]).decode("utf-16-le")
        self.End = self.Start + units(value)
        if doc.TrackRevisions:
            doc.Revisions.items.append(NS(Type=9, Author="tester", Range=NS(Text=value)))
        refresh(doc)
        if doc.fail == "edit":
            raise error()

    def Information(self, kind, /):
        return {3: 1, 12: self.doc.in_table, 13: 2, 16: 3}[kind]


def refresh(doc):
    doc.Content.Start, doc.Content.End = 0, units(doc.Content.Text)
    rows, start = [], 0
    for line in doc.Content.Text.split("\r")[:-1]:
        end = start + units(line) + 1
        rows.append(NS(Range=EditRange(doc, start, end)))
        start = end
    doc.Paragraphs = Collection(rows)


def edits(o, text="before target after\rnext paragraph\r"):
    word(o)
    doc = o.doc
    doc.story, doc.style, doc.list_string, doc.mixed = 1, "Body Text", "1.", False
    doc.range_objects = {}
    doc.native_fields = (*doc.native_fields, "Revisions", "TrackRevisions")
    doc.Range = lambda start, end: EditRange(doc, start, end)
    doc.Content.Text = text
    doc.Bookmarks, doc.Hyperlinks, doc.ContentControls, doc.Fields = Collection(), Collection(), Collection(), Collection()
    refresh(doc)
    return doc


class Comment:
    def __init__(self, doc, value, start=7, end=13, parent=None, author="Reviewer", date=None):
        self.doc, self.Author, self.Initial, self.Done = doc, author, "RV", False
        self.Date = date or datetime(2026, 1, 2, 3, 4, 5)
        self.Scope = doc.Range(start, end)
        self.Range = NS(Text=value)
        self.Ancestor, self.Replies = parent, Replies(self)

    def Delete(self, /):
        self.doc.calls.append(("delete_comment",))
        self.doc.Comments.items.remove(self)


class Replies(Collection):
    def __init__(self, parent):
        super().__init__()
        self.parent = parent

    def Add(self, scope, text, /):
        parent = self.parent
        reply = Comment(parent.doc, text, scope.Start, scope.End, parent)
        self.items.append(reply)
        parent.doc.Comments.items.append(reply)
        parent.doc.calls.append(("reply", text))
        return reply
