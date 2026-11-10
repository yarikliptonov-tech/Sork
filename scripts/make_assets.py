"""Графики, схемы и иконки для документа и презентации -> build/assets/."""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import Arc, Rectangle
from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).resolve().parent.parent / "build" / "assets"
OUT.mkdir(parents=True, exist_ok=True)
INK = "#0F172A"
MUTED = "#64748B"
plt.rcParams.update({"font.family": "DejaVu Sans", "text.color": INK, "axes.labelcolor": INK,
                     "xtick.color": INK, "ytick.color": INK})

EMOJI = {
    "kitchen": "🍳", "light": "💡", "sound": "🔊", "comm": "📡", "home": "🏠", "shield": "🛡",
    "money": "💰", "target": "🎯", "idea": "💭", "pot": "🍲", "magnet": "🧲", "snow": "❄",
    "spoon": "🥄", "drop": "💧", "thread": "🧵", "phone": "📱", "tv": "📺", "nophone": "📵",
    "wifi": "📶", "globe": "🌐", "kettle": "🫖", "warn": "⚠", "headphones": "🎧", "check": "✅",
    "mirror": "🪞", "lens": "🔍", "ear": "👂", "bolt": "⚡", "books": "📚", "storm": "⛈",
    "micro": "🔬", "timer": "⏱", "wave": "🌊", "mute": "🔇", "door": "🚪", "calc": "🧮",
}


def emoji_icons():
    font = ImageFont.truetype("/usr/share/fonts/truetype/noto/NotoColorEmoji.ttf", 109)
    for name, ch in EMOJI.items():
        img = Image.new("RGBA", (200, 200), (0, 0, 0, 0))
        ImageDraw.Draw(img).text((100, 100), ch, font=font, embedded_color=True, anchor="mm")
        box = img.getbbox()
        assert box, f"emoji {name} не отрисовался"
        img = img.crop(box)
        side = max(img.size) + 8
        sq = Image.new("RGBA", (side, side), (0, 0, 0, 0))
        sq.paste(img, ((side - img.width) // 2, (side - img.height) // 2))
        sq.save(OUT / f"e_{name}.png")


def _clean(ax):
    for s in ax.spines.values():
        s.set_visible(False)
    ax.tick_params(length=0)


def _save(fig, name):
    fig.savefig(OUT / name, dpi=200, transparent=True)
    plt.close(fig)


def boil():
    # Давление по барометрической формуле, температура кипения по уравнению Антуана.
    def t_boil(p_kpa):
        mm = p_kpa * 7.50062
        a, b, c = (8.07131, 1730.63, 233.426) if mm <= 760 else (8.14019, 1810.94, 244.485)
        return b / (a - np.log10(mm)) - c

    def p_alt(h):
        return 101.325 * (1 - 2.25577e-5 * h) ** 5.25588

    items = [("Эверест, 8849 м", t_boil(p_alt(8849)), "#60A5FA"),
             ("Эльбрус, 5642 м", t_boil(p_alt(5642)), "#22D3EE"),
             ("Уровень моря", t_boil(101.325), "#FB923C"),
             ("Скороварка", t_boil(200.0), "#EF4444")]  # около 2 бар под крышкой
    vals = [round(v) for _, v, _ in items]
    assert vals == [70, 81, 100, 120], vals
    fig, ax = plt.subplots(figsize=(6.4, 3.4))
    ax.barh([i[0] for i in items], vals, color=[i[2] for i in items], height=0.62)
    for i, v in enumerate(vals):
        ax.text(v + 2, i, f"{v} °C", va="center", fontsize=20, fontweight="bold")
    ax.set_xlim(0, 166)
    ax.set_xticks([])
    ax.tick_params(axis="y", labelsize=17)
    _clean(ax)
    fig.tight_layout()
    _save(fig, "boil.png")


def lamps():
    fig, ax = plt.subplots(figsize=(5.0, 3.2))
    ax.bar(["Лампа\nнакаливания", "Светодиодная\nлампа"], [75, 10], color=["#F97316", "#22C55E"], width=0.55)
    ax.text(0, 77, "75 Вт", ha="center", fontsize=18, fontweight="bold")
    ax.text(1, 12, "10 Вт", ha="center", fontsize=18, fontweight="bold")
    ax.set_ylim(0, 95)
    ax.set_yticks([])
    ax.tick_params(axis="x", labelsize=19)
    _clean(ax)
    fig.tight_layout()
    _save(fig, "lamps.png")


def rgb():
    n = 700
    y, x = np.mgrid[0:n, 0:n]
    img = np.zeros((n, n, 4), dtype=np.uint8)
    r = 0.30 * n
    centers = [(0.5 * n, 0.34 * n), (0.34 * n, 0.63 * n), (0.66 * n, 0.63 * n)]
    for ch, (cx, cy) in enumerate(centers):
        mask = (x - cx) ** 2 + (y - cy) ** 2 <= r ** 2
        img[..., ch][mask] = 255
        img[..., 3][mask] = 255
    Image.fromarray(img, "RGBA").resize((500, 500), Image.LANCZOS).save(OUT / "rgb.png")


def optics():
    fig, axes = plt.subplots(1, 2, figsize=(9, 4))
    for ax in axes:
        ax.set_xlim(-2, 2)
        ax.set_ylim(-1.9, 2.1)
        ax.set_aspect("equal")
        ax.axis("off")
    ray = dict(arrowprops=dict(arrowstyle="-|>", color="#F59E0B", lw=3.2, mutation_scale=22))

    ax = axes[0]
    ax.add_patch(Rectangle((-1.9, -0.22), 3.8, 0.22, facecolor="#CBD5E1", edgecolor="#64748B", hatch="////", lw=1))
    ax.plot([-1.9, 1.9], [0, 0], color="#334155", lw=3)
    ax.plot([0, 0], [0, 1.9], color=MUTED, lw=1.5, ls="--")
    ax.annotate("", xy=(0, 0), xytext=(-1.5, 1.5), **ray)
    ax.annotate("", xy=(1.5, 1.5), xytext=(0, 0), **ray)
    ax.add_patch(Arc((0, 0), 1.3, 1.3, theta1=90, theta2=135, color=INK, lw=1.6))
    ax.add_patch(Arc((0, 0), 1.5, 1.5, theta1=45, theta2=90, color=INK, lw=1.6))
    ax.text(-0.36, 0.82, "α", fontsize=20, ha="center")
    ax.text(0.4, 0.9, "β", fontsize=20, ha="center")
    ax.text(0, -0.62, "зеркало", fontsize=17, ha="center", color=MUTED)
    ax.text(0, -1.45, "Отражение: α = β", fontsize=18, ha="center", fontweight="bold")

    ax = axes[1]
    ax.add_patch(Rectangle((-1.9, -1.15), 3.8, 1.15, facecolor="#BAE6FD", edgecolor="none"))
    ax.plot([-1.9, 1.9], [0, 0], color="#0284C7", lw=2.5)
    ax.plot([0, 0], [-1.15, 1.9], color=MUTED, lw=1.5, ls="--")
    a1 = np.arctan(1.2 / 1.5)
    a2 = np.arcsin(np.sin(a1) / 1.33)
    ax.annotate("", xy=(0, 0), xytext=(-1.2, 1.5), **ray)
    ax.annotate("", xy=(1.1 * np.tan(a2), -1.1), xytext=(0, 0), **ray)
    ax.plot([0, 1.1 * np.tan(a1)], [0, -1.1], color="#F59E0B", lw=1.8, ls=":")
    ax.text(1.85, 0.16, "воздух", fontsize=17, ha="right", color=MUTED)
    ax.text(1.85, -0.36, "вода", fontsize=17, ha="right", color="#0369A1")
    ax.text(0, -1.45, "Преломление: луч «ломается»", fontsize=18, ha="center", fontweight="bold")
    fig.tight_layout()
    _save(fig, "optics.png")


def sound_speed():
    fig, ax = plt.subplots(figsize=(6.4, 4.2))
    names, vals = ["Воздух", "Вода", "Сталь"], [340, 1500, 5000]
    ax.bar(names, vals, color=["#C4B5FD", "#8B5CF6", "#5B21B6"], width=0.6)
    for i, v in enumerate(vals):
        ax.text(i, v + 120, f"{v} м/с", ha="center", fontsize=19, fontweight="bold")
    ax.set_ylim(0, 5900)
    ax.set_yticks([])
    ax.tick_params(axis="x", labelsize=16)
    _clean(ax)
    fig.tight_layout()
    _save(fig, "sound_speed.png")


def db():
    fig, ax = plt.subplots(figsize=(7, 4))
    items = [("Шёпот", 30, "#22C55E"), ("Разговор", 60, "#A3E635"), ("Пылесос", 75, "#FBBF24"), ("Блендер", 90, "#EF4444")]
    ax.barh([i[0] for i in items], [i[1] for i in items], color=[i[2] for i in items], height=0.6)
    for i, (_, v, _) in enumerate(items):
        ax.text(v - 1.5, i, f"{v} дБ", va="center", ha="right", fontsize=17, fontweight="bold")
    ax.axvline(80, color="#DC2626", lw=2, ls="--", ymin=0, ymax=0.9)
    ax.text(80, 3.72, "80 дБ — долгий шум вреден", color="#DC2626", fontsize=15, ha="center", fontweight="bold")
    ax.set_xlim(0, 100)
    ax.set_ylim(-0.5, 4.0)
    ax.set_xticks([])
    ax.tick_params(axis="y", labelsize=17)
    _clean(ax)
    fig.tight_layout()
    _save(fig, "db.png")


def spectrum():
    fig, ax = plt.subplots(figsize=(11, 3.4))
    bands = [("Радиоволны", 2.3, "#2563EB", "FM-радио, ТВ\nλ ≈ 3 м"),
             ("Микроволны", 2.5, "#0891B2", "Wi-Fi, телефон,\nмикроволновка\nλ ≈ 12 см"),
             ("Инфракрасные", 2.3, "#B91C1C", "пульт, тепло\nот батареи"),
             ("Видимый\nсвет", 1.5, None, "лампы,\nэкран"),
             ("Ультрафиолет", 2.3, "#7C3AED", "солнце,\nкварцевая\nлампа"),
             ("Рентген", 1.5, "#475569", "снимок\nв поликлинике")]
    rainbow = LinearSegmentedColormap.from_list("vis", ["#EF4444", "#F97316", "#FACC15", "#22C55E", "#3B82F6", "#8B5CF6"])
    x = 0.0
    for name, w, color, example in bands:
        if color:
            ax.add_patch(Rectangle((x, 0), w - 0.06, 1, facecolor=color, edgecolor="none"))
        else:
            ax.imshow(np.linspace(0, 1, 256).reshape(1, -1), extent=(x, x + w - 0.06, 0, 1), cmap=rainbow, aspect="auto")
        ax.text(x + w / 2, 0.5, name, ha="center", va="center", color="white", fontsize=15, fontweight="bold")
        ax.text(x + w / 2, -0.18, example, ha="center", va="top", fontsize=14, color=INK, linespacing=1.3)
        x += w
    ax.annotate("", xy=(x - 0.1, 1.42), xytext=(0.05, 1.42), arrowprops=dict(arrowstyle="-|>", color=MUTED, lw=2, mutation_scale=18))
    ax.text(x / 2, 1.56, "длина волны уменьшается, частота растёт", ha="center", va="bottom", fontsize=14, color=MUTED)
    ax.set_xlim(0, x)
    ax.set_ylim(-1.75, 2.1)
    ax.axis("off")
    fig.subplots_adjust(left=0.01, right=0.99, top=0.99, bottom=0.01)
    _save(fig, "spectrum.png")


if __name__ == "__main__":
    emoji_icons()
    for f in (boil, lamps, rgb, optics, sound_speed, db, spectrum):
        f()
    print("assets:", len(list(OUT.glob("*.png"))))
