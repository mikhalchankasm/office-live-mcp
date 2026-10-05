"""In-memory protocol registry, shared by installer and protocol unit tests."""

import copy

from office_live.protocol import tree


class FakeRegistry:
    def __init__(self):
        self.snapshot = None
        self.calls = []

    def read(self):
        self.calls.append("read")
        return copy.deepcopy(self.snapshot)

    def write(self, command):
        self.calls.append("write")
        self.snapshot = tree(command)

    def delete(self):
        self.calls.append("delete")
        self.snapshot = None
