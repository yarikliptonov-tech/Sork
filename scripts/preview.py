"""Контактный лист из картинок для визуальной проверки: preview.py out.png cols cell_width files..."""
import sys

from PIL import Image


def sheet(out, cols, cell_w, files):
    ims = []
    for f in files:
        im = Image.open(f).convert("RGBA")
        im = im.resize((cell_w, round(im.height * cell_w / im.width)), Image.LANCZOS)
        bg = Image.new("RGBA", im.size, "white")
        bg.alpha_composite(im)
        ims.append(bg.convert("RGB"))
    rows = [ims[i:i + cols] for i in range(0, len(ims), cols)]
    gap = 12
    height = sum(max(i.height for i in r) + gap for r in rows) + gap
    canvas = Image.new("RGB", (cols * (cell_w + gap) + gap, height), "#94A3B8")
    y = gap
    for r in rows:
        for j, im in enumerate(r):
            canvas.paste(im, (gap + j * (cell_w + gap), y))
        y += max(i.height for i in r) + gap
    canvas.save(out)
    print(out, canvas.size)


if __name__ == "__main__":
    sheet(sys.argv[1], int(sys.argv[2]), int(sys.argv[3]), sys.argv[4:])
