"""Собирает out/Физика_в_быту_текст.docx: текст выступления, поделённый на четырёх участников."""
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

import content as C

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "build" / "assets"
OUT = ROOT / "out" / "Физика_в_быту_текст.docx"
INK, MUTED = "0F172A", "64748B"
WORDS_PER_MIN = 120
PAGE_W = 17.0  # ширина текста, см

doc = Document()
for sect in doc.sections:
    sect.page_width, sect.page_height = Cm(21), Cm(29.7)
    sect.left_margin = sect.right_margin = Cm(2)
    sect.top_margin = sect.bottom_margin = Cm(2)
normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(12.5)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.2


def rgb(h):
    return RGBColor.from_string(h)


def run(p, t, size=None, color=None, bold=False, italic=False, highlight=False):
    r = p.add_run(t)
    r.bold, r.italic = bold, italic
    if size:
        r.font.size = Pt(size)
    if color:
        r.font.color.rgb = rgb(color)
    if highlight:
        r.font.highlight_color = WD_COLOR_INDEX.YELLOW
    return r


def para(t="", size=None, color=None, bold=False, italic=False, align=None, before=0, after=6, keep=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(before), Pt(after)
    p.paragraph_format.keep_with_next = keep
    if align:
        p.alignment = align
    if t:
        run(p, t, size, color, bold, italic)
    return p


def shade(cell, fill):
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def pad(cell, top=140, bottom=110, side=200):
    """Внутренние поля ячейки в твипах (после shade: порядок элементов в tcPr важен)."""
    mar = OxmlElement("w:tcMar")
    for edge, v in (("top", top), ("left", side), ("bottom", bottom), ("right", side)):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:w"), str(v))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    cell._tc.get_or_add_tcPr().append(mar)


def fix_widths(table, widths):
    table.autofit = False
    for col, w in zip(table.columns, widths):
        col.width = Cm(w)
    for row in table.rows:
        for c, w in zip(row.cells, widths):
            c.width = Cm(w)


def block(fill, width=PAGE_W):
    """Цветная плашка во всю ширину: таблица 1×1 с заливкой."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    fix_widths(t, [width])
    cell = t.rows[0].cells[0]
    shade(cell, fill)
    pad(cell)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(0), Pt(4)
    return cell, p


def cell_para(cell, before=0, after=6):
    p = cell.add_paragraph()
    p.paragraph_format.space_before, p.paragraph_format.space_after = Pt(before), Pt(after)
    return p


def minutes(n_words):
    m = n_words / WORDS_PER_MIN
    return f"{m:.1f}".replace(".", ",")


def cell_text(cell, t, **kw):
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.1
    return run(p, t, **kw)


# ---------- Титульный лист ----------
para("Школа: ______________________________________________", size=12, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER, after=70)
cell, p = block("0B1437")
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(20)
run(p, "ПРОЕКТНАЯ РАБОТА ПО ФИЗИКЕ", size=13, color="93C5FD", bold=True)
p = cell_para(cell, after=4)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, "ФИЗИКА В БЫТУ", size=44, color="FFFFFF", bold=True)
p = cell_para(cell, after=26)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run(p, C.SUBTITLE, size=16, color="CBD5E1")
t = doc.add_table(rows=1, cols=4)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
fix_widths(t, [PAGE_W / 4] * 4)
for c, key in zip(t.rows[0].cells, ["kitchen", "light", "sound", "comm"]):
    shade(c, C.SECTIONS[key]["color"])
    c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.paragraphs[0].paragraph_format.space_before = c.paragraphs[0].paragraph_format.space_after = Pt(5)
    run(c.paragraphs[0], C.SECTIONS[key]["label"], size=12, color="FFFFFF", bold=True)
para("Текст выступления для четырёх участников", size=15, color=INK, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, before=34, after=40)
para("Выполнили ученики ______ класса:", size=12.5, after=4)
for i in C.SPEAKERS:
    p = para(after=4)
    run(p, f"{C.SPEAKERS[i]['name']}: ", bold=True, color=C.SPEAKERS[i]["color"])
    run(p, "____________________________________")
para("Руководитель: ____________________________________", size=12.5, before=10, after=60)
para("2026", size=12.5, color=MUTED, align=WD_ALIGN_PARAGRAPH.CENTER)
doc.add_page_break()

# ---------- Кто что говорит ----------
para("КТО ЧТО ГОВОРИТ", size=20, color=INK, bold=True, after=4)
total = C.words(sorted(C.SLIDES))
para(f"Весь текст — {total} слов, это около {round(total / WORDS_PER_MIN)} минут спокойной речи. "
     "Часть каждого участника начинается с новой страницы — её можно распечатать отдельно. "
     "Номера слайдов совпадают с презентацией.", color=MUTED, size=11.5, after=10)
t = doc.add_table(rows=1, cols=5)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
widths = [2.8, 7.4, 2.8, 1.8, 2.2]
for c, h in zip(t.rows[0].cells, ["Кто", "Что рассказывает", "Слайды", "Слов", "Минут"]):
    shade(c, "0B1437")
    cell_text(c, h, size=11.5, color="FFFFFF", bold=True)
for i, sp in C.SPEAKERS.items():
    nums = C.speaker_slides(i)
    n_words = C.words(nums)
    cells = t.add_row().cells
    shade(cells[0], sp["color"])
    cell_text(cells[0], sp["name"], size=11.5, color="FFFFFF", bold=True)
    cell_text(cells[1], sp["parts"], size=11.5)
    cell_text(cells[2], C.ranges(nums), size=11.5)
    cell_text(cells[3], str(n_words), size=11.5, bold=True)
    cell_text(cells[4], "≈ " + minutes(n_words), size=11.5)
fix_widths(t, widths)
para(after=8)
cell, p = block("FEF9C3")
run(p, "Обратите внимание. ", bold=True, size=11.5)
run(p, "Жёлтым", size=11.5, highlight=True)
run(p, " в тексте выделены числа-примеры из опыта с чайником. Проведите опыт сами и впишите свои "
       "измерения — в тексте и на слайде 7. Под описанием опытов есть подсказки, что подготовить.", size=11.5)
para(after=10)
para("СТРУКТУРА РАБОТЫ", size=14, color=INK, bold=True, after=4)
for part in C.PARTS:
    sp = C.SPEAKERS[part["speaker"]]
    p = para(after=2)
    run(p, "■ ", color=C.SECTIONS[part["section"]]["color"], size=12)
    run(p, part["toc"], size=12)
    run(p, f"  —  {sp['name']}, слайды {C.ranges(part['slides'])}", size=11, color=MUTED)
p = para(after=2)
run(p, "■ ", color=MUTED, size=12)
run(p, "Список литературы", size=12)
p = para(after=2)
run(p, "■ ", color=MUTED, size=12)
run(p, "Приложение. Памятка «Физика помогает дома»", size=12)

# ---------- Части по участникам ----------
for part in C.PARTS:
    sp, sec = C.SPEAKERS[part["speaker"]], C.SECTIONS[part["section"]]
    if part.get("same_page"):
        para(after=6)
    else:
        doc.add_page_break()
    cell, p = block(sec["color"])
    run(p, f"{sp['name'].upper()}  ·  слайды {C.ranges(part['slides'])}  ·  ≈ {minutes(C.words(part['slides']))} мин",
        size=10.5, color="FFFFFF", bold=True)
    p = cell_para(cell, after=6)
    run(p, part["heading"], size=19, color="FFFFFF", bold=True)
    para(after=2)
    for n in part["slides"]:
        d = C.SLIDES[n]
        p = para(before=8, after=4, keep=True)
        run(p, f"Слайд {n}  ", size=11, color=MUTED, bold=True)
        run(p, d["doc_title"], size=14.5, color=sec["color"], bold=True)
        for sp_text in d["speech"]:
            p = para(after=6)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            for chunk, marked in C.split_marks(sp_text):
                run(p, chunk, highlight=marked, bold=marked)
        if d.get("figure"):
            p = para(before=4, after=6, align=WD_ALIGN_PARAGRAPH.CENTER)
            wide = d["figure"] in ("spectrum.png", "optics.png")
            p.add_run().add_picture(str(ASSETS / d["figure"]), width=Cm(16.5 if wide else 10.5))
        if d.get("tip"):
            cell, p = block(sec["tint"])
            run(p, d["tip"], size=11, color=INK, italic=True)
            para(after=2)

# ---------- Список литературы ----------
doc.add_page_break()
cell, p = block("0B1437")
run(p, "СПИСОК ЛИТЕРАТУРЫ", size=19, color="FFFFFF", bold=True)
para(after=4)
for i, src in enumerate(C.SOURCES, 1):
    para(f"{i}. {src}", after=4)
para("Год издания учебников укажите по своим экземплярам.", size=11, color=MUTED, italic=True, before=4, after=16)

# ---------- Приложение: памятка ----------
cell, p = block(C.SECTIONS["final"]["color"])
run(p, "ПРИЛОЖЕНИЕ", size=10.5, color="FFFFFF", bold=True)
p = cell_para(cell)
run(p, "Памятка «Физика помогает дома»", size=19, color="FFFFFF", bold=True)
para(after=4)
t = doc.add_table(rows=0, cols=2)
t.style = "Table Grid"
t.alignment = WD_TABLE_ALIGNMENT.CENTER
colors = ["EA580C", "EA580C", "D97706", "EA580C", "7C3AED", "0D9488"]
for (_ico, tip, why), color in zip(C.TIPS, colors):
    cells = t.add_row().cells
    shade(cells[0], color)
    cell_text(cells[0], tip, size=12, color="FFFFFF", bold=True)
    cell_text(cells[1], why, size=12)
fix_widths(t, [PAGE_W / 2] * 2)

OUT.parent.mkdir(exist_ok=True)
doc.save(OUT)
print("saved", OUT.name, "| слов:", {i: C.words(C.speaker_slides(i)) for i in C.SPEAKERS}, "всего", total)
