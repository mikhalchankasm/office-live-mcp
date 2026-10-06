"""Read-only threaded comment metadata and stable branch guards for Word."""

import hashlib
import json

import pywintypes

from . import com, wd_common as wd
from .errors import ToolError
from .util import clean_word_text, truncate


def optional(obj, name, unavailable):
    try:
        return getattr(obj, name)
    except (AttributeError, pywintypes.com_error) as exc:
        if com.is_busy(exc) or com.is_dead(exc):
            raise
        unavailable.add(name)
        return None


def date_of(comment, unavailable):
    value = optional(comment, "Date", unavailable)
    if value is None:
        return None
    try:
        return value.isoformat()
    except AttributeError:
        unavailable.add("Date")
        return None


def signature(comment, unavailable):
    return [str(comment.Author), date_of(comment, unavailable), str(comment.Range.Text), int(comment.Scope.Start)]


def root_of(comment, unavailable):
    root, seen = comment, set()
    for _ in range(30):
        token = json.dumps(signature(root, unavailable), ensure_ascii=False)
        if token in seen:
            break
        seen.add(token)
        ancestor = optional(root, "Ancestor", unavailable)
        if ancestor is None:
            break
        if signature(ancestor, unavailable) == signature(root, unavailable):
            break
        root = ancestor
    return root


def thread_key(comment, unavailable=None):
    unavailable = unavailable if unavailable is not None else set()
    root = root_of(comment, unavailable)
    return hashlib.sha256(json.dumps(signature(root, unavailable), ensure_ascii=False).encode("utf-8")).hexdigest()[:20]


def check_thread(doc, index, expected):
    if expected and (not index or thread_key(doc.Comments(int(index))) != expected):
        raise ToolError(f"Комментарии изменились: индекс {index} теперь относится к другой ветке; перечитайте список")


def list_comments(doc, offset=0, limit=None, max_chars=None, only_open=False, context_chars=120):
    # None preserves legacy callers: up to 300, full comment text. New pagination opts into 50/1000.
    paginated = limit is not None or max_chars is not None or offset != 0 or only_open or context_chars != 120
    limit = (50 if paginated else 300) if limit is None else limit
    max_chars = (1000 if paginated else None) if max_chars is None else max_chars
    if type(offset) is not int or offset < 0 or type(limit) is not int or not 1 <= limit <= 300:
        raise ToolError("offset must be nonnegative; limit must be from 1 to 300.")
    if max_chars is not None and (type(max_chars) is not int or not 1 <= max_chars <= 100000):
        raise ToolError("max_chars must be from 1 to 100000.")
    if type(context_chars) is not int or not 0 <= context_chars <= 10000:
        raise ToolError("context_chars must be from 0 to 10000.")
    unavailable, selected = set(), []
    n = int(doc.Comments.Count)
    for i in range(1, n + 1):
        comment = doc.Comments(i)
        done = optional(comment, "Done", unavailable)
        if only_open and done is not None and bool(done):
            continue
        selected.append((i, comment, None if done is None else bool(done)))
    index_by_signature = None

    def index_of(target):
        nonlocal index_by_signature
        if index_by_signature is None:  # one pass over the collection instead of a scan per reply
            index_by_signature = {}
            for k in range(1, n + 1):
                index_by_signature.setdefault(json.dumps(signature(doc.Comments(k), unavailable), ensure_ascii=False), k)
        return index_by_signature.get(json.dumps(signature(target, unavailable), ensure_ascii=False))

    items = []
    for i, comment, done in selected[offset:offset + limit]:
        scope = comment.Scope
        parent = optional(comment, "Ancestor", unavailable)
        is_reply = None if "Ancestor" in unavailable else bool(parent is not None and signature(parent, unavailable) != signature(comment, unavailable))
        parent_index = None
        if is_reply:
            parent_index = index_of(parent)
        value = clean_word_text(comment.Range.Text)
        shortened = max_chars is not None and len(value) > max_chars
        location = wd.range_location(doc, scope)
        start, end = int(scope.Start), int(scope.End)
        main = int(scope.StoryType) == 1
        before = doc.Range(max(int(doc.Content.Start), start - context_chars), start).Text if main and context_chars else ""
        after = doc.Range(end, min(int(doc.Content.End), end + context_chars)).Text if main and context_chars else ""
        replies = None
        collection = optional(comment, "Replies", unavailable)
        if is_reply is False:
            if collection is not None:
                replies = []
                for j in range(1, int(collection.Count) + 1):
                    reply = collection.Item(j)
                    reply_text = clean_word_text(reply.Range.Text)
                    reply_index = index_of(reply)
                    clipped = max_chars is not None and len(reply_text) > max_chars
                    replies.append({"index": reply_index, "author": reply.Author, "date": date_of(reply, unavailable),
                                    "text": reply_text[:max_chars] if clipped else reply_text, "truncated": clipped})
                    shortened |= clipped
        items.append({"index": i, "author": comment.Author, "anchored_text": truncate(clean_word_text(scope.Text), 120)[0],
                      "comment": value[:max_chars] if max_chars is not None else value, "resolved": done, "paragraph": location["paragraph"],
                      "date": date_of(comment, unavailable), "initials": optional(comment, "Initial", unavailable),
                      "is_reply": is_reply, "parent_index": parent_index, "replies": replies,
                      "thread_key": thread_key(comment, unavailable), "context_before": str(before)[-context_chars:] if context_chars else "",
                      "context_after": str(after)[:context_chars], "location": location,
                      "link": wd.range_link(doc, location["paragraph"], scope) if main else None, "truncated": bool(shortened)})
    total = len(selected)
    return {"document": doc.Name, "count": n, "comments": items, "total": total, "returned": len(items),
            "next_offset": offset + len(items) if offset + len(items) < total else None,
            "unavailable_properties": sorted(unavailable)}
