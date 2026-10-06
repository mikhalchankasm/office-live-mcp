"""Генератор графики README: баннер, карточки преимуществ, схема — EN/RU × светлая/тёмная тема.

Запуск: python docs/assets/build_assets.py  (пишет *.svg рядом с собой). Текст правится здесь, а не в SVG.
"""

from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
FONT = "Segoe UI, system-ui, -apple-system, Helvetica, Arial, sans-serif"

THEMES = {
    "light": {"bg1": "#f6f8fa", "bg2": "#eaf1fb", "surface": "#ffffff", "border": "#d0d7de", "ink": "#1f2328", "muted": "#57606a",
              "grid": "#e6ebf1", "excel": "#217346", "word": "#2b579a", "agent": "#6e56cf", "agent2": "#8a73e6", "ins": "#2b579a",
              "del": "#cf222e", "hl": "#dafbe1", "chip": "#ffffff", "shadow": "#8c959f"},
    "dark": {"bg1": "#161b22", "bg2": "#0d1117", "surface": "#1c2129", "border": "#30363d", "ink": "#f0f6fc", "muted": "#9198a1",
             "grid": "#2a313a", "excel": "#2f8f58", "word": "#3d6fbf", "agent": "#8a73e6", "agent2": "#a691ff", "ins": "#79a7ef",
             "del": "#ff7b72", "hl": "#1f3b2a", "chip": "#161b22", "shadow": "#000000"},
}

TEXT = {
    "en": {
        "tagline": "Your AI agent in the Excel and Word files you already have open",
        "pills": ["Any MCP agent", "Live Excel & Word", "Preview & undo", "Free · MIT"],
        "chat": ["Make an act per row", "of this register — as", "tracked changes"],
        "features": [
            ("chat", "agent", "Your agent, your model", ["Keep your instructions and workflow.", "Any model your MCP client supports."]),
            ("pipeline", "excel", "Excel → Word → PDF", ["A document per spreadsheet row from", "your template, previewed first."]),
            ("shield", "word", "Access you control", ["Read-only mode, allowed folders, tool", "sets — enforced by the server."]),
            ("eye", "excel", "Changes you can see", ["Works in open documents, unsaved", "edits included. Results appear at once."]),
            ("undo", "word", "Review and undo", ["Word edits as tracked revisions.", "Journal and undo with conflict checks."]),
            ("code", "agent", "Free and open", ["No server subscription, MIT source.", "One installer, no admin rights."]),
        ],
        "how": [("Your AI agent", ["Claude Code · Cursor · Codex", "ZCode · VS Code · any MCP client"]),
                ("Office Live MCP", ["runs locally · read-only mode", "allowed folders · journal & undo"]),
                ("Excel and Word", ["the files you already have open", "changes appear on screen"])],
        "how_title": "How it works",
    },
    "ru": {
        "tagline": "Ваш ИИ-агент в уже открытых файлах Excel и Word",
        "pills": ["Любой MCP-агент", "Excel и Word вживую", "Предпросмотр и отмена", "Бесплатно · MIT"],
        "chat": ["Сделай акт на каждую", "строку реестра — через", "исправления"],
        "features": [
            ("chat", "agent", "Свой агент и модель", ["Ваши инструкции и порядок работы.", "Любая модель вашего MCP-клиента."]),
            ("pipeline", "excel", "Excel → Word → PDF", ["Документ на каждую строку реестра", "по шаблону, с предпросмотром."]),
            ("shield", "word", "Доступ под контролем", ["Только чтение, разрешённые папки,", "наборы — проверяет сервер."]),
            ("eye", "excel", "Правки сразу видны", ["Работа в открытых документах,", "включая несохранённые правки."]),
            ("undo", "word", "Просмотр и отмена", ["Правки Word через исправления.", "Журнал и отмена с проверкой."]),
            ("code", "agent", "Бесплатно и открыто", ["Без подписки на сервер, код MIT.", "Один установщик, без прав админа."]),
        ],
        "how": [("Ваш ИИ-агент", ["Claude Code · Cursor · Codex", "ZCode · VS Code · любой MCP-клиент"]),
                ("Office Live MCP", ["локально · режим только чтения", "разрешённые папки · журнал и отмена"]),
                ("Excel и Word", ["уже открытые файлы", "правки сразу на экране"])],
        "how_title": "Как это работает",
    },
}


def t(x, y, s, size, fill, weight=400, anchor="start"):
    return (f'<text x="{x}" y="{y}" font-family="{FONT}" font-size="{size}" font-weight="{weight}" fill="{fill}" '
            f'text-anchor="{anchor}">{escape(s)}</text>')


def width(s, size, weight=400):
    """Грубая оценка ширины строки (для плашек): средняя ширина знака Segoe UI ≈ 0.55em, жирного ≈ 0.6em."""
    return len(s) * size * (0.6 if weight >= 600 else 0.55)


def icon(kind, cx, cy, color):
    """Простые линейные иконки 40×40 в едином стиле (обводка 2.6, скруглённые концы)."""
    s = f'fill="none" stroke="{color}" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"'
    x, y = cx - 14, cy - 14
    shapes = {
        "chat": f'<path {s} d="M{x+2} {y+5} h24 a3 3 0 0 1 3 3 v12 a3 3 0 0 1 -3 3 h-12 l-7 6 v-6 h-5 a3 3 0 0 1 -3 -3 v-12 a3 3 0 0 1 3 -3 z"/>'
                f'<circle cx="{x+9}" cy="{y+14}" r="1.6" fill="{color}"/><circle cx="{x+15}" cy="{y+14}" r="1.6" fill="{color}"/><circle cx="{x+21}" cy="{y+14}" r="1.6" fill="{color}"/>',
        "pipeline": f'<rect {s} x="{x}" y="{y+4}" width="11" height="14" rx="2"/><path {s} d="M{x} {y+9} h11 M{x} {y+13.5} h11 M{x+5.5} {y+4} v14"/>'
                    f'<path {s} d="M{x+13} {y+11} h5 m-2.5 -2.5 l2.5 2.5 l-2.5 2.5"/><path {s} d="M{x+20} {y+4} h6 l4 4 v14 h-10 z M{x+22.5} {y+12} h5 M{x+22.5} {y+16} h5"/>',
        "shield": f'<path {s} d="M{cx} {y+1} l11 4 v8 c0 7 -5 12 -11 14 c-6 -2 -11 -7 -11 -14 v-8 z"/><path {s} d="M{cx-5} {cy+1} l3.5 3.5 l7 -7"/>',
        "eye": f'<path {s} d="M{x} {cy} c4 -7 9 -10 14 -10 s10 3 14 10 c-4 7 -9 10 -14 10 s-10 -3 -14 -10 z"/><circle {s} cx="{cx}" cy="{cy}" r="4.5"/>',
        "undo": f'<path {s} d="M{x+4} {y+9} h14 a7 7 0 0 1 0 14 h-9"/><path {s} d="M{x+9} {y+3} l-6 6 l6 6"/>',
        "code": f'<path {s} d="M{x+7} {y+6} l-7 8 l7 8 M{x+21} {y+6} l7 8 l-7 8 M{x+17} {y+2} l-6 24"/>',
    }
    return shapes[kind]


def frame(w, h, c, rx=20):
    return (f'<defs><linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="{c["bg1"]}"/>'
            f'<stop offset="1" stop-color="{c["bg2"]}"/></linearGradient>'
            f'<filter id="sh" x="-10%" y="-10%" width="120%" height="130%"><feDropShadow dx="0" dy="6" stdDeviation="9" flood-color="{c["shadow"]}" flood-opacity="0.18"/></filter></defs>'
            f'<rect width="{w}" height="{h}" rx="{rx}" fill="url(#bg)"/><rect x="0.5" y="0.5" width="{w-1}" height="{h-1}" rx="{rx}" fill="none" stroke="{c["border"]}"/>')


def window(x, y, w, h, c, accent, title):
    return (f'<g filter="url(#sh)"><rect x="{x}" y="{y}" width="{w}" height="{h}" rx="12" fill="{c["surface"]}" stroke="{c["border"]}"/></g>'
            f'<path d="M{x} {y+12} a12 12 0 0 1 12 -12 h{w-24} a12 12 0 0 1 12 12 v18 h-{w} z" fill="{accent}"/>'
            + t(x + 16, y + 21, title, 13, "#ffffff", 600))


def hero(lang, theme):
    c, T = THEMES[theme], TEXT[lang]
    W, H = 1200, 420
    out = [frame(W, H, c)]
    out.append(t(64, 150, "Office Live MCP", 60, c["ink"], 700))
    tag = T["tagline"]
    words, lines, cur = tag.split(), [], ""
    for w_ in words:  # перенос подзаголовка в две строки по ширине ~500px
        if width(cur + " " + w_, 25) > 500 and cur:
            lines.append(cur)
            cur = w_
        else:
            cur = (cur + " " + w_).strip()
    lines.append(cur)
    for i, line in enumerate(lines):
        out.append(t(66, 196 + i * 34, line, 25, c["muted"]))
    py = 196 + len(lines) * 34 + 26
    px = 66
    colors = [c["agent"], c["excel"], c["word"], c["muted"]]
    for n, (label, col) in enumerate(zip(T["pills"], colors, strict=True)):
        pw = width(label, 15, 600) * 0.9 + 26
        if n == 2:  # две плашки в ряд: ровная сетка и свободное место под репликой агента справа
            px, py = 66, py + 44
        out.append(f'<rect x="{px:.0f}" y="{py}" width="{pw:.0f}" height="32" rx="16" fill="{c["chip"]}" stroke="{c["border"]}"/>')
        out.append(t(px + pw / 2, py + 21, label, 15, col, 600, "middle"))
        px += pw + 10
    # Excel window
    ex, ey = 640, 70
    out.append(window(ex, ey, 300, 210, c, c["excel"], "Register.xlsx"))
    gx, gy, cw, ch = ex + 14, ey + 44, 68, 26
    for r in range(6):
        for k in range(4):
            fill = c["hl"] if (r in (1, 2) and k == 3) else c["surface"]
            out.append(f'<rect x="{gx + k*cw}" y="{gy + r*ch}" width="{cw}" height="{ch}" fill="{fill}" stroke="{c["grid"]}"/>')
    for k, head in enumerate(["Code", "Client", "Sum", "Act"]):
        out.append(t(gx + k * cw + 8, gy + 18, head, 12, c["muted"], 600))
    rows = [["007", "Alfa", "1 200", "✓"], ["0012", "Beta", "860", "✓"], ["0031", "Gamma", "2 400", "…"], ["0040", "Delta", "515", "…"], ["0057", "Omega", "990", ""]]
    for r, row in enumerate(rows, start=1):
        for k, val in enumerate(row):
            out.append(t(gx + k * cw + 8, gy + r * ch + 18, val, 12, c["ink"]))
    # Word window
    wx, wy = 880, 170
    out.append(window(wx, wy, 270, 220, c, c["word"], "Act_0031.docx"))
    for ln, frac in [(0, 0.9), (1, 0.75), (3, 0.85), (4, 0.6)]:
        out.append(f'<rect x="{wx+20}" y="{wy+50+ln*22}" width="{(230*frac):.0f}" height="7" rx="3.5" fill="{c["grid"]}"/>')
    out.append(t(wx + 20, wy + 101, "Sum: ", 13, c["ink"], 600))
    out.append(f'<text x="{wx+60}" y="{wy+101}" font-family="{FONT}" font-size="13" fill="{c["del"]}" text-decoration="line-through">2 300</text>')
    out.append(f'<text x="{wx+102}" y="{wy+101}" font-family="{FONT}" font-size="13" fill="{c["ins"]}" text-decoration="underline">2 400</text>')
    out.append(f'<rect x="{wx+12}" y="{wy+86}" width="3" height="20" rx="1.5" fill="{c["ins"]}"/>')
    for i, frac in enumerate([0.8, 0.5]):
        out.append(f'<rect x="{wx+20}" y="{wy+160+i*22}" width="{(230*frac):.0f}" height="7" rx="3.5" fill="{c["grid"]}"/>')
    # agent chat bubble
    bx, by = 600, 296
    out.append(f'<g filter="url(#sh)"><rect x="{bx}" y="{by}" width="250" height="92" rx="16" fill="{c["agent"]}"/></g>')
    out.append(f'<path d="M{bx+36} {by+92} l-10 16 l26 -16 z" fill="{c["agent"]}"/>')
    for i, line in enumerate(T["chat"]):
        out.append(t(bx + 20, by + 30 + i * 22, line, 15, "#ffffff", 600 if i == 0 else 400))
    # dashed links agent → windows
    out.append(f'<path d="M{bx+250} {by+30} C {bx+300} {by+20}, {wx-30} {wy+150}, {wx} {wy+150}" fill="none" stroke="{c["agent2"]}" stroke-width="3" stroke-dasharray="2 8" stroke-linecap="round"/>')
    desc = escape(f"Office Live MCP — {T['tagline']}")
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}" role="img" aria-label="{desc}">'
            + "".join(out) + "</svg>\n")


def features(lang, theme):
    c, T = THEMES[theme], TEXT[lang]
    W, H, cols, cw, chh, gap = 1200, 400, 3, 376, 176, 16
    out = [f'<defs><filter id="sh" x="-10%" y="-10%" width="120%" height="130%"><feDropShadow dx="0" dy="4" stdDeviation="7" flood-color="{c["shadow"]}" flood-opacity="0.12"/></filter></defs>']
    for i, (ic, col, title, lines) in enumerate(T["features"]):
        x = 8 + (i % cols) * (cw + gap)
        y = 8 + (i // cols) * (chh + gap)
        accent = c[col]
        out.append(f'<g filter="url(#sh)"><rect x="{x}" y="{y}" width="{cw}" height="{chh}" rx="16" fill="{c["surface"]}" stroke="{c["border"]}"/></g>')
        out.append(f'<rect x="{x+24}" y="{y+24}" width="52" height="52" rx="14" fill="{accent}" fill-opacity="0.12"/>')
        out.append(icon(ic, x + 50, y + 50, accent))
        out.append(t(x + 92, y + 58, title, 21, c["ink"], 700))
        for k, line in enumerate(lines):
            out.append(t(x + 24, y + 116 + k * 26, line, 16, c["muted"]))
    alt = escape("; ".join(f"{f[2]}: {' '.join(f[3])}" for f in T["features"]))
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}" role="img" aria-label="{alt}">'
            + "".join(out) + "</svg>\n")


def how(lang, theme):
    c, T = THEMES[theme], TEXT[lang]
    W, H = 1200, 250
    out = [frame(W, H, c)]
    out.append(t(W / 2, 46, T["how_title"], 22, c["ink"], 700, "middle"))
    boxes = [(60, c["agent"]), (440, c["agent"]), (820, c["excel"])]
    for i, ((title, lines), (x, accent)) in enumerate(zip(T["how"], boxes, strict=True)):
        y = 76
        main = i == 1
        stroke = accent if main else c["border"]
        out.append(f'<g filter="url(#sh)"><rect x="{x}" y="{y}" width="320" height="136" rx="16" fill="{c["surface"]}" stroke="{stroke}" stroke-width="{2 if main else 1}"/></g>')
        if i == 2:
            out.append(f'<rect x="{x+24}" y="{y+24}" width="26" height="26" rx="7" fill="{c["excel"]}"/>' + t(x + 37, y + 43, "X", 16, "#fff", 700, "middle"))
            out.append(f'<rect x="{x+56}" y="{y+24}" width="26" height="26" rx="7" fill="{c["word"]}"/>' + t(x + 69, y + 43, "W", 16, "#fff", 700, "middle"))
            tx = x + 94
        else:
            out.append(f'<rect x="{x+20}" y="{y+18}" width="38" height="38" rx="11" fill="{accent}" fill-opacity="0.14"/>')
            out.append(icon("chat" if i == 0 else "shield", x + 39, y + 37, accent))
            tx = x + 70
        out.append(t(tx, y + 44, title, 20, c["ink"], 700))
        for k, line in enumerate(lines):
            out.append(t(x + 24, y + 86 + k * 24, line, 15, c["muted"]))
        if i < 2:
            ax = x + 320
            out.append(f'<path d="M{ax+8} {y+68} h44" stroke="{c["agent2"]}" stroke-width="3" stroke-linecap="round"/>'
                       f'<path d="M{ax+44} {y+60} l10 8 l-10 8" fill="none" stroke="{c["agent2"]}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>')
    alt = escape(" → ".join(f"{a}: {' · '.join(b)}" for a, b in T["how"]))
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}" role="img" aria-label="{alt}">'
            + "".join(out) + "</svg>\n")


def main():
    for lang in TEXT:
        suffix = "" if lang == "en" else "-ru"
        for theme in THEMES:
            (HERE / f"hero{suffix}-{theme}.svg").write_text(hero(lang, theme), encoding="utf-8")
            (HERE / f"features{suffix}-{theme}.svg").write_text(features(lang, theme), encoding="utf-8")
            (HERE / f"how{suffix}-{theme}.svg").write_text(how(lang, theme), encoding="utf-8")
    print("written:", sorted(p.name for p in HERE.glob("*.svg")))


if __name__ == "__main__":
    main()
