"""Password arguments are ephemeral: redact before logging or retaining history."""

import json

MASK = "<redacted>"


def redact(arguments):
    if isinstance(arguments, dict):
        return {k: MASK if "password" in str(k).casefold() else redact(v) for k, v in arguments.items()}
    if isinstance(arguments, (list, tuple)):
        return [redact(v) for v in arguments]
    return arguments


def safe_text(text, arguments):
    secrets = []

    def collect(value):
        if isinstance(value, dict):
            for key, item in value.items():
                if "password" in str(key).casefold() and isinstance(item, str) and item:
                    secrets.extend((item, repr(item)[1:-1], json.dumps(item, ensure_ascii=True)[1:-1]))
                else:
                    collect(item)
        elif isinstance(value, (list, tuple)):
            for item in value:
                collect(item)

    collect(arguments)
    text = str(text)
    for secret in sorted(set(secrets), key=len, reverse=True):
        text = text.replace(secret, MASK)
    return text
