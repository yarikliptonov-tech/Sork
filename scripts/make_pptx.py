"""Собирает out/Физика_в_быту_презентация.pptx (19 слайдов, текст выступления — в заметках)."""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

import re

import content as C

ROOT = Path(__file__).resolve().parent.parent
ASSETS = ROOT / "build" / "assets"
OUT = ROOT / "out" / "Физика_в_быту_презентация.pptx"

W, H = 13.333, 7.5
X0, CW = 0.75, 11.85  # левый край и ширина области содержимого
DARK, INK, BODY, MUTED = "0B1437", "0F172A", "334155", "64748B"
BG, WHITE, BORDER, YELLOW = "F8FAFC", "FFFFFF", "E2E8F0", "FDE047"
FONT = "Calibri"
ALIGN = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
ANCHOR = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)
BLANK = prs.slide_layouts[6]


def rgb(h):
    return RGBColor.from_string(h)


_UNIT = re.compile(r"(\d) (°C|%|Вт|ГГц|Гц|дБ|км/с|м/с|км|см|Дж|л|м|с)(?![А-Яа-яЁё\w])")


def nbsp(t):
    """Неразрывный пробел между числом и единицей, чтобы «20 °C» не рвалось по строкам."""
    return _UNIT.sub("\\1\u00a0\\2", t)


def box(s, x, y, w, h, fill, line=None, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08):
    sh = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(fill)
    if line:
        sh.line.color.rgb = rgb(line)
        sh.line.width = Pt(1.25)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        sh.adjustments[0] = radius
    return sh


def text(s, x, y, w, h, paras, size=18, color=INK, bold=False, align="l", anchor="t", spacing=1.08, after=6):
    """paras — строка или список абзацев; **так** выделяется полужирный."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = ANCHOR[anchor]
    for i, p in enumerate([paras] if isinstance(paras, str) else paras):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = ALIGN[align]
        para.line_spacing = spacing
        para.space_after = Pt(after)
        for j, chunk in enumerate(nbsp(p).split("**")):
            if not chunk:
                continue
            r = para.add_run()
            r.text = chunk
            r.font.name = FONT
            r.font.size = Pt(size)
            r.font.bold = bold or j % 2 == 1
            r.font.color.rgb = rgb(color)
    return tb


def icon(s, name, x, y, size):
    s.shapes.add_picture(str(ASSETS / f"e_{name}.png"), Inches(x), Inches(y), Inches(size), Inches(size))


def badge(s, name, x, y, d, fill):
    """Иконка в цветном круге."""
    box(s, x, y, d, d, fill, shape=MSO_SHAPE.OVAL)
    pad = d * 0.2
    icon(s, name, x + pad, y + pad, d - 2 * pad)


def picture(s, name, x, y, w):
    return s.shapes.add_picture(str(ASSETS / name), Inches(x), Inches(y), width=Inches(w))


def strip(s, y, msg, color, h=0.85, size=20, x=X0, w=CW, fg=WHITE):
    box(s, x, y, w, h, color, radius=0.25)
    text(s, x + 0.35, y, w - 0.7, h, msg, size=size, color=fg, bold=True, align="c", anchor="m")


def notes(s, n):
    d = C.SLIDES[n]
    who = C.SPEAKERS[d["speaker"]]["name"]
    body = "\n\n".join(C.plain(p) for p in d["speech"])
    extra = f"\n\n[{d['tip']}]" if d.get("tip") else ""
    s.notes_slide.notes_text_frame.text = f"ГОВОРИТ: {who.upper()}\n\n{body}{extra}"


def base(n, title):
    """Светлый слайд: цветная полоса раздела, плашка, заголовок, номер."""
    d = C.SLIDES[n]
    sec = C.SECTIONS[d["section"]]
    s = prs.slides.add_slide(BLANK)
    box(s, 0, 0, W, H, BG, shape=MSO_SHAPE.RECTANGLE)
    box(s, 11.75, -1.15, 2.6, 2.6, sec["tint"], shape=MSO_SHAPE.OVAL)
    box(s, 0, 0, 0.32, H, sec["color"], shape=MSO_SHAPE.RECTANGLE)
    label = sec["label"]
    box(s, X0, 0.45, 0.13 * len(label) + 0.55, 0.42, sec["color"], radius=0.5)
    text(s, X0, 0.45, 0.13 * len(label) + 0.55, 0.42, label, size=13, color=WHITE, bold=True, align="c", anchor="m")
    text(s, X0, 1.0, CW - 0.6, 0.8, title, size=34, bold=True, anchor="m")
    text(s, X0, 7.02, 5, 0.3, "Физика в быту", size=11, color=MUTED)
    text(s, W - 1.75, 7.02, 1.0, 0.3, str(n), size=11, color=MUTED, align="r")
    notes(s, n)
    return s, sec


def card(s, x, y, w, h, ico, head, body, sec, tag=None, body_size=17):
    box(s, x, y, w, h, WHITE, line=BORDER, radius=0.06)
    badge(s, ico, x + 0.3, y + 0.3, 0.95, sec["tint"])
    text(s, x + 0.3, y + 1.45, w - 0.6, 0.45, head, size=21, bold=True, anchor="m")
    text(s, x + 0.3, y + 2.0, w - 0.6, h - 2.7, body, size=body_size, color=BODY)
    if tag:
        box(s, x + 0.3, y + h - 0.7, w - 0.6, 0.42, sec["tint"], radius=0.5)
        text(s, x + 0.3, y + h - 0.7, w - 0.6, 0.42, tag, size=13, color=sec["color"], bold=True, align="c", anchor="m")


def mini(s, x, y, w, h, head, body, sec, body_size=15):
    box(s, x, y, w, h, WHITE, line=BORDER, radius=0.08)
    text(s, x + 0.3, y + 0.22, w - 0.6, 0.4, head, size=19, color=sec["color"], bold=True, anchor="m")
    text(s, x + 0.3, y + 0.72, w - 0.6, h - 0.85, body, size=body_size, color=BODY)


def fact(s, x, y, w, h, ico, msg, sec, size=17):
    box(s, x, y, w, h, WHITE, line=BORDER, radius=0.14)
    d = min(0.8, h - 0.3)
    badge(s, ico, x + 0.25, y + (h - d) / 2, d, sec["tint"])
    text(s, x + 0.3 + d + 0.25, y, w - d - 0.85, h, msg, size=size, color=BODY, anchor="m")


def exp_card(s, x, y, w, h, ico, head, did, saw, why, sec):
    box(s, x, y, w, h, WHITE, line=BORDER, radius=0.05)
    badge(s, ico, x + 0.3, y + 0.3, 0.9, sec["tint"])
    text(s, x + 1.4, y + 0.3, w - 1.7, 0.9, head, size=23, bold=True, anchor="m")
    for i, (label, body) in enumerate([("ЧТО ДЕЛАЛИ", did), ("ЧТО УВИДЕЛИ", saw), ("ПОЧЕМУ", why)]):
        yy = y + 1.45 + i * 1.08
        box(s, x + 0.3, yy + 0.04, 0.07, 0.9, sec["color"], shape=MSO_SHAPE.RECTANGLE)
        text(s, x + 0.55, yy, w - 0.85, 0.28, label, size=12, color=sec["color"], bold=True)
        text(s, x + 0.55, yy + 0.3, w - 0.85, 0.7, body, size=17, color=BODY)


def dark_slide(n):
    s = prs.slides.add_slide(BLANK)
    box(s, 0, 0, W, H, DARK, shape=MSO_SHAPE.RECTANGLE)
    notes(s, n)
    return s


def areas_grid(s, x, y):
    """Четыре области проекта: 2×2 цветных круга с подписями."""
    items = [("kitchen", "Кухня"), ("light", "Свет"), ("sound", "Звук"), ("comm", "Связь")]
    for i, (key, label) in enumerate(items):
        cx, cy = x + (i % 2) * 2.35, y + (i // 2) * 2.55
        badge(s, key, cx, cy, 1.75, C.SECTIONS[key]["color"])
        text(s, cx - 0.3, cy + 1.85, 2.35, 0.4, label, size=18, color=WHITE, bold=True, align="c")


# ---------- 1. Титул ----------
s = dark_slide(1)
box(s, -1.6, 5.3, 4.2, 4.2, "13225A", shape=MSO_SHAPE.OVAL)
box(s, 6.2, -2.0, 3.6, 3.6, "13225A", shape=MSO_SHAPE.OVAL)
box(s, 0.9, 0.8, 4.3, 0.46, "2563EB", radius=0.5)
text(s, 0.9, 0.8, 4.3, 0.46, "ПРОЕКТНАЯ РАБОТА ПО ФИЗИКЕ", size=14, color=WHITE, bold=True, align="c", anchor="m")
text(s, 0.9, 1.55, 7.2, 2.7, ["ФИЗИКА", "В БЫТУ"], size=88, color=WHITE, bold=True, spacing=0.9, after=0)
box(s, 0.9, 4.45, 1.4, 0.09, "F97316", shape=MSO_SHAPE.RECTANGLE)
text(s, 0.9, 4.7, 7.0, 0.5, C.SUBTITLE, size=24, color="CBD5E1")
text(s, 0.9, 5.85, 7.2, 0.9,
     ["Выполнили: (имена четырёх участников), (класс)", "Руководитель: (ФИО учителя)"],
     size=15, color="94A3B8", after=3)
areas_grid(s, 8.45, 1.15)

# ---------- 2. Актуальность ----------
s, sec = base(2, "Почему мы выбрали эту тему")
cw = (CW - 0.6) / 3
for i, (ico, head, body) in enumerate([
    ("home", "Физика рядом", "Чайник, лампа, колонка, роутер — мы пользуемся ими каждый день, но редко знаем, как они работают."),
    ("shield", "Безопасность", "Почему нельзя ставить металл в микроволновку? Чем опасен громкий звук в наушниках?"),
    ("money", "Экономия", "Какая лампа тратит меньше? Сколько воды наливать в чайник? Ответы даёт физика."),
]):
    card(s, X0 + i * (cw + 0.3), 2.0, cw, 3.55, ico, head, body, sec)
strip(s, 5.85, "Физика помогает не просто пользоваться вещами, а понимать их", sec["color"])

# ---------- 3. Цель и задачи ----------
s, sec = base(3, "Цель и задачи проекта")
box(s, X0, 2.0, 4.3, 4.6, sec["color"], radius=0.06)
badge(s, "target", X0 + 0.35, 2.35, 1.0, "1D4ED8")
text(s, X0 + 0.35, 3.6, 3.6, 0.35, "ЦЕЛЬ", size=14, color=sec["tint"], bold=True)
text(s, X0 + 0.35, 4.0, 3.6, 2.4, "Показать, что привычные явления дома объясняются законами физики из курса 7–9 классов",
     size=22, color=WHITE, bold=True)
for i, t in enumerate([
    "Изучить литературу и выбрать явления: кухня, свет, звук, связь",
    "Объяснить каждое явление законами физики",
    "Провести простые опыты из подручных предметов",
    "Сделать выводы и составить памятку с полезными советами",
]):
    y = 2.0 + i * 1.2
    box(s, 5.35, y, 7.25, 1.0, WHITE, line=BORDER, radius=0.14)
    box(s, 5.6, y + 0.2, 0.6, 0.6, sec["color"], shape=MSO_SHAPE.OVAL)
    text(s, 5.6, y + 0.2, 0.6, 0.6, str(i + 1), size=20, color=WHITE, bold=True, align="c", anchor="m")
    text(s, 6.45, y, 5.9, 1.0, t, size=18, color=BODY, anchor="m")

# ---------- 4. Гипотеза ----------
s, sec = base(4, "Гипотеза и методы")
box(s, X0, 2.0, CW, 2.5, sec["color"], radius=0.08)
badge(s, "idea", X0 + 0.4, 2.65, 1.2, "1D4ED8")
text(s, 2.7, 2.3, 9.5, 0.35, "ГИПОТЕЗА", size=14, color=sec["tint"], bold=True)
text(s, 2.7, 2.7, 9.6, 1.7,
     "Если знать законы школьной физики, можно объяснить работу большинства бытовых приборов — "
     "и пользоваться ими безопаснее и экономнее.", size=25, color=WHITE, bold=True)
for i, (head, body) in enumerate([
    ("ОБЪЕКТ", "Физические явления в обычной квартире"),
    ("ПРЕДМЕТ", "Объяснение этих явлений законами физики"),
    ("МЕТОДЫ", "Литература, наблюдение, опыт, измерение, расчёт"),
]):
    mini(s, X0 + i * (cw + 0.3), 4.8, cw, 1.8, head, body, sec, body_size=18)

# ---------- 5. Кипение ----------
s, sec = base(5, "Когда закипает вода?")
for i, (ico, t) in enumerate([
    ("kettle", "**100 °C** — только при нормальном атмосферном давлении"),
    ("pot", "**Скороварка:** давление вдвое выше — вода кипит при 120 °C"),
    ("snow", "**Эльбрус:** давление вдвое ниже — вода кипит всего при 81 °C"),
]):
    fact(s, X0, 2.0 + i * 1.25, 5.6, 1.1, ico, t, sec)
box(s, 6.65, 2.0, 5.95, 3.6, WHITE, line=BORDER, radius=0.05)
text(s, 6.95, 2.15, 5.4, 0.35, "Температура кипения воды", size=15, color=sec["color"], bold=True)
picture(s, "boil.png", 6.85, 2.55, 5.55)
strip(s, 5.85, "Чем выше давление, тем выше температура кипения", sec["color"])

# ---------- 6. Приборы ----------
s, sec = base(6, "Как работают кухонные приборы")
for i, (ico, head, body, tag) in enumerate([
    ("wave", "Микроволновая печь", "Волны частотой 2,45 ГГц раскачивают молекулы воды в еде — их энергия превращается в тепло.", "электромагнитные волны"),
    ("magnet", "Индукционная плита", "Переменное магнитное поле создаёт токи в дне кастрюли. Греется сама посуда, а не плита.", "электромагнитная индукция"),
    ("snow", "Холодильник", "Хладагент испаряется внутри и забирает тепло, а отдаёт его снаружи — поэтому задняя стенка тёплая.", "испарение и конденсация"),
]):
    card(s, X0 + i * (cw + 0.3), 2.0, cw, 4.6, ico, head, body, sec, tag=tag, body_size=18)

# ---------- 7. Опыт: чайник ----------
s, sec = base(7, "Опыт на кухне: сколько энергии тратит чайник?")
box(s, X0, 2.0, 3.55, 3.6, WHITE, line=BORDER, radius=0.06)
text(s, X0 + 0.3, 2.2, 3.0, 0.3, "ЧТО ДЕЛАЛИ", size=13, color=sec["color"], bold=True)
for i, t in enumerate(["Налили в чайник 1 л воды температурой 20 °C", "Засекли время до закипания", "Сравнили время с расчётом"]):
    y = 2.7 + i * 0.95
    box(s, X0 + 0.3, y + 0.05, 0.5, 0.5, sec["color"], shape=MSO_SHAPE.OVAL)
    text(s, X0 + 0.3, y + 0.05, 0.5, 0.5, str(i + 1), size=17, color=WHITE, bold=True, align="c", anchor="m")
    text(s, X0 + 1.0, y - 0.1, 2.35, 0.8, t, size=16, color=BODY, anchor="m")
box(s, 4.55, 2.0, 4.35, 3.6, INK, radius=0.06)
text(s, 4.85, 2.2, 3.8, 0.3, "РАСЧЁТ", size=13, color="FDBA74", bold=True)
text(s, 4.85, 2.6, 3.8, 0.7, "Q = c · m · Δt", size=34, color=WHITE, bold=True, anchor="m")
text(s, 4.85, 3.5, 3.8, 1.9,
     ["Q = 4200 · 1 · 80 = **336 000 Дж**", "t = Q : P = 336 000 : 2000 = **168 с**", "КПД = Q : (P · t)"],
     size=18, color="E2E8F0", after=12)
box(s, 9.15, 2.0, 3.45, 3.6, sec["tint"], radius=0.06)
text(s, 9.45, 2.2, 2.9, 0.3, "РЕЗУЛЬТАТ", size=13, color=sec["color"], bold=True)
text(s, 9.45, 2.6, 2.9, 0.9, ["По расчёту — **168 с**", "В опыте — **195 с**"], size=19, color=INK, after=4)
text(s, 9.45, 3.6, 2.9, 0.9, "КПД ≈ 86 %", size=38, color=sec["color"], bold=True, anchor="m")
# Наклейка-напоминание: участники удаляют её, когда впишут свои измерения.
box(s, 9.45, 4.7, 2.85, 0.65, YELLOW, radius=0.2)
text(s, 9.55, 4.7, 2.65, 0.65, "Пример: впишите числа своего опыта", size=12, color=INK, bold=True, align="c", anchor="m")
strip(s, 5.85, "Часть энергии уходит на нагрев чайника и воздуха. Кипятите столько воды, сколько нужно!", sec["color"], size=18)

# ---------- 8. Отражение и преломление ----------
s, sec = base(8, "Отражение и преломление света")
box(s, X0, 2.0, 7.2, 3.6, WHITE, line=BORDER, radius=0.05)
picture(s, "optics.png", X0 + 0.1, 2.2, 7.0)
mini(s, 8.25, 2.0, 4.35, 2.2, "Зеркало", "Угол падения равен углу отражения. Изображение — «за зеркалом», на таком же расстоянии.", sec, body_size=17)
mini(s, 8.25, 4.4, 4.35, 2.3, "Линзы", "Свет меняет направление на границе воздуха и стекла. Так работают очки, лупа, камера и наш глаз.", sec, body_size=17)
strip(s, 5.85, "Скорость света — 300 000 км/с", sec["color"], w=7.2)

# ---------- 9. Лампы и экран ----------
s, sec = base(9, "Лампы и экран смартфона")
box(s, X0, 2.0, 5.75, 4.6, WHITE, line=BORDER, radius=0.05)
text(s, X0 + 0.3, 2.2, 5.1, 0.4, "Какая лампа экономнее?", size=20, color=sec["color"], bold=True, anchor="m")
picture(s, "lamps.png", X0 + 0.55, 2.65, 4.6)
text(s, X0 + 0.3, 5.65, 5.15, 0.9, "Светят одинаково, но светодиодная тратит **в 7 раз меньше**. Лампа накаливания превращает в свет лишь около 5 % энергии.",
     size=15, color=BODY)
box(s, 6.85, 2.0, 5.75, 4.6, WHITE, line=BORDER, radius=0.05)
text(s, 7.15, 2.2, 5.1, 0.4, "Из чего состоит цвет на экране?", size=20, color=sec["color"], bold=True, anchor="m")
box(s, 7.15, 2.8, 2.9, 2.9, INK, radius=0.12)
picture(s, "rgb.png", 7.3, 2.95, 2.6)
text(s, 10.3, 2.9, 2.1, 2.8, ["Пиксель — это три огонька:", "**красный, зелёный и синий.**", "Все три вместе дают белый цвет."], size=17, color=BODY, after=10)
text(s, 7.15, 5.9, 5.15, 0.6, "Смешивая три цвета, экран показывает миллионы оттенков", size=15, color=BODY)

# ---------- 10. Опыты со светом ----------
s, sec = base(10, "Опыты со светом")
ew = (CW - 0.4) / 2
exp_card(s, X0, 2.0, ew, 4.75, "spoon", "«Сломанная ложка»",
         "Опустили ложку в стакан с водой и посмотрели сбоку.",
         "На границе воды и воздуха ложка выглядит переломленной.",
         "Свет преломляется, когда выходит из воды в воздух.", sec)
exp_card(s, X0 + ew + 0.4, 2.0, ew, 4.75, "drop", "«Капля-лупа»",
         "Капнули воду на экран телефона с белой картинкой.",
         "Сквозь каплю видны красные, зелёные и синие точки.",
         "Капля — выпуклая линза: она увеличивает пиксели экрана.", sec)

# ---------- 11. Звук ----------
s, sec = base(11, "Звук — волна, которая бежит по веществу")
for i, (ico, t) in enumerate([
    ("sound", "**Звук — это колебания**, которые передаются по воздуху, воде и твёрдым телам"),
    ("mute", "**В вакууме звука нет:** там нечему колебаться"),
    ("storm", "**Гроза:** 3 секунды между молнией и громом — это примерно 1 км"),
]):
    fact(s, X0, 2.0 + i * 1.55, 5.6, 1.4, ico, t, sec)
box(s, 6.65, 2.0, 5.95, 4.5, WHITE, line=BORDER, radius=0.05)
text(s, 6.95, 2.15, 5.4, 0.4, "Скорость звука", size=20, color=sec["color"], bold=True, anchor="m")
picture(s, "sound_speed.png", 6.9, 2.65, 5.5)

# ---------- 12. Громкость ----------
s, sec = base(12, "Громкость, слух и эхо")
box(s, X0, 2.0, 7.0, 4.6, WHITE, line=BORDER, radius=0.05)
text(s, X0 + 0.3, 2.15, 6.4, 0.4, "Громкость в децибелах", size=20, color=sec["color"], bold=True, anchor="m")
picture(s, "db.png", X0 + 0.2, 2.65, 6.6)
mini(s, 8.05, 2.0, 4.55, 1.8, "Что мы слышим", "Звуки частотой от 20 до 20 000 Гц. Чем больше частота, тем выше звук.", sec, body_size=17)
mini(s, 8.05, 3.95, 4.55, 1.8, "Откуда эхо", "Голые стены отражают звук. Ковры, шторы и мебель его поглощают.", sec, body_size=17)
strip(s, 5.9, "Наушники — не на полную громкость!", sec["color"], h=0.7, size=17, x=8.05, w=4.55)

# ---------- 13. Опыты со звуком ----------
s, sec = base(13, "Опыты со звуком")
exp_card(s, X0, 2.0, ew, 4.75, "thread", "«Нитяной телефон»",
         "Соединили два бумажных стаканчика натянутой ниткой длиной 5 м.",
         "Шёпот хорошо слышно. Ослабили нитку — звук пропал.",
         "Колебания донышка бегут по нитке: твёрдое тело проводит звук.", sec)
exp_card(s, X0 + ew + 0.4, 2.0, ew, 4.75, "phone", "«Телефон в стакане»",
         "Поставили играющий телефон в пустой стакан.",
         "Звук стал заметно громче.",
         "Стакан работает как рупор: направляет звук в одну сторону.", sec)

# ---------- 14. Электромагнитные волны ----------
s, sec = base(14, "Одна семья электромагнитных волн")
box(s, X0, 2.0, CW, 3.6, WHITE, line=BORDER, radius=0.05)
picture(s, "spectrum.png", X0 + 0.35, 2.08, CW - 0.7)
for i, t in enumerate(["Все летят со скоростью света", "Отличаются частотой и длиной волны", "Свет и сигнал телефона — «родственники»"]):
    box(s, X0 + i * (cw + 0.3), 5.85, cw, 0.85, sec["tint"], radius=0.25)
    text(s, X0 + i * (cw + 0.3) + 0.15, 5.85, cw - 0.3, 0.85, t, size=16, color=sec["color"], bold=True, align="c", anchor="m")

# ---------- 15. Роутер, микроволновка, оптоволокно ----------
s, sec = base(15, "Роутер, микроволновка и оптоволокно")
for i, (ico, head, body) in enumerate([
    ("wifi", "Почти одна частота", "Роутер и микроволновка работают около 2,4 ГГц. Но мощность роутера — 0,1 Вт, а печи — 800 Вт."),
    ("door", "Сетка в дверце печи", "Отверстия — пара миллиметров, а волна — 12 см: она не проходит. Свет проходит, и мы видим еду."),
    ("globe", "Оптоволокно", "Интернет приходит в дом вспышками света по стеклянной нити: свет отражается от её стенок."),
]):
    card(s, X0 + i * (cw + 0.3), 2.0, cw, 3.65, ico, head, body, sec)
strip(s, 5.9, "Металл отражает радиоволны — поэтому в лифте пропадает связь", sec["color"], h=0.8)

# ---------- 16. Опыты со связью ----------
s, sec = base(16, "Опыты со связью")
exp_card(s, X0, 2.0, ew, 4.75, "tv", "«Невидимый луч пульта»",
         "Навели пульт на камеру телефона и нажали кнопку.",
         "Глаз ничего не видит, а на экране мигает огонёк.",
         "Это инфракрасное излучение: камера его замечает, а глаз — нет.", sec)
exp_card(s, X0 + ew + 0.4, 2.0, ew, 4.75, "nophone", "«Телефон в фольге»",
         "Завернули телефон в несколько слоёв фольги и позвонили на него.",
         "Абонент «вне зоны доступа». Развернули — связь вернулась.",
         "Металл не пропускает радиоволны. Это клетка Фарадея.", sec)

# ---------- 17. Памятка ----------
s, sec = base(17, "Памятка «Физика помогает дома»")
for i, (ico, tip, _why) in enumerate(C.TIPS):
    x, y = X0 + (i % 3) * (cw + 0.3), 2.0 + (i // 3) * 2.4
    box(s, x, y, cw, 2.15, WHITE, line=BORDER, radius=0.08)
    badge(s, ico, x + 0.3, y + 0.25, 0.85, sec["tint"])
    text(s, x + 0.3, y + 1.2, cw - 0.6, 0.85, tip, size=17, bold=True, anchor="m")

# ---------- 18. Выводы ----------
s, sec = base(18, "Выводы")
for i, t in enumerate([
    "Изучили **4 области** быта и более **10 явлений** и приборов",
    "Каждое объяснили законами физики 7–9 классов",
    "Провели **7 опытов** с обычными домашними предметами",
    "Составили памятку с полезными советами",
]):
    fact(s, X0, 2.0 + i * 1.2, 6.6, 1.0, "check", t, sec, size=18)
box(s, 7.65, 2.0, 4.95, 4.6, sec["color"], radius=0.06)
badge(s, "micro", 7.95, 2.3, 1.1, "BE185D")
text(s, 7.95, 3.6, 4.4, 1.1, ["ГИПОТЕЗА", "ПОДТВЕРДИЛАСЬ"], size=28, color=WHITE, bold=True, spacing=0.95, after=0)
text(s, 7.95, 4.85, 4.4, 1.6, "Знание физики помогает понимать, как работает дом, — и пользоваться им безопаснее и экономнее.",
     size=18, color="FCE7F3")

# ---------- 19. Спасибо ----------
s = dark_slide(19)
box(s, -1.6, 5.3, 4.2, 4.2, "13225A", shape=MSO_SHAPE.OVAL)
box(s, 6.2, -2.0, 3.6, 3.6, "13225A", shape=MSO_SHAPE.OVAL)
text(s, 0.9, 1.0, 7.3, 2.3, ["Спасибо", "за внимание!"], size=66, color=WHITE, bold=True, spacing=0.92, after=0)
box(s, 0.9, 3.4, 1.4, 0.09, "F97316", shape=MSO_SHAPE.RECTANGLE)
text(s, 0.9, 3.6, 7.0, 0.5, "Готовы ответить на ваши вопросы", size=24, color="CBD5E1")
text(s, 0.9, 4.55, 7.0, 0.3, "ИСТОЧНИКИ", size=13, color="93C5FD", bold=True)
text(s, 0.9, 4.9, 7.2, 2.2, [f"{i}. {src}" for i, src in enumerate(C.SOURCES, 1)], size=13, color="CBD5E1", after=2)
areas_grid(s, 8.45, 1.15)

assert len(prs.slides) == len(C.SLIDES) == 19
OUT.parent.mkdir(exist_ok=True)
prs.save(OUT)
print("saved", OUT.name, len(prs.slides), "slides")
