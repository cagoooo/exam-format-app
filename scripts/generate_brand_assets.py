from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
ASSET_DIR = ROOT / "static" / "assets"
TITLE = "考卷格式自動校正系統"
SUBTITLE = "固定版面、頁數檢查、Word 與 PDF 輸出"
SCHOOL = "桃園市龍潭區石門國民小學"


COLORS = {
    "ink": "#17211b",
    "muted": "#66736d",
    "paper": "#fbfcfa",
    "panel": "#ffffff",
    "accent": "#246b54",
    "accent_dark": "#174c3a",
    "line": "#d9ded8",
    "gold": "#d2a24c",
    "red": "#b94a48",
}


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = [
        "C:/Windows/Fonts/msjhbd.ttc" if bold else "C:/Windows/Fonts/msjh.ttc",
        "C:/Windows/Fonts/kaiu.ttf",
        "C:/Windows/Fonts/mingliu.ttc",
    ]
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return ImageFont.truetype(str(path), size)
    return ImageFont.load_default()


def rounded_rectangle(draw: ImageDraw.ImageDraw, xy, radius, fill, outline=None, width=1):
    draw.rounded_rectangle(xy, radius=radius, fill=fill, outline=outline, width=width)


def draw_doc_icon(draw: ImageDraw.ImageDraw, x: int, y: int, scale: float = 1.0):
    w, h = int(92 * scale), int(120 * scale)
    fold = int(24 * scale)
    rounded_rectangle(draw, (x, y, x + w, y + h), int(10 * scale), COLORS["panel"], COLORS["accent_dark"], max(2, int(3 * scale)))
    draw.polygon([(x + w - fold, y), (x + w, y + fold), (x + w - fold, y + fold)], fill="#e8f0ea", outline=COLORS["accent_dark"])
    for i in range(4):
        yy = y + int((38 + i * 17) * scale)
        draw.line((x + int(18 * scale), yy, x + w - int(18 * scale), yy), fill=COLORS["accent"], width=max(2, int(3 * scale)))
    draw.rectangle((x + int(18 * scale), y + int(92 * scale), x + int(52 * scale), y + int(104 * scale)), fill=COLORS["gold"])


def make_icon(size: int, maskable: bool = False) -> Image.Image:
    img = Image.new("RGBA", (size, size), COLORS["accent"])
    draw = ImageDraw.Draw(img)
    margin = int(size * (0.12 if maskable else 0.07))
    rounded_rectangle(draw, (margin, margin, size - margin, size - margin), int(size * 0.18), COLORS["paper"])
    draw_doc_icon(draw, int(size * 0.27), int(size * 0.20), size / 260)
    badge_r = int(size * 0.16)
    cx, cy = int(size * 0.65), int(size * 0.67)
    draw.ellipse((cx - badge_r, cy - badge_r, cx + badge_r, cy + badge_r), fill=COLORS["accent"], outline=COLORS["accent_dark"], width=max(2, size // 36))
    check = [
        (cx - int(size * 0.08), cy),
        (cx - int(size * 0.025), cy + int(size * 0.055)),
        (cx + int(size * 0.095), cy - int(size * 0.075)),
    ]
    draw.line(check, fill="white", width=max(4, size // 18), joint="curve")
    return img


def make_og() -> Image.Image:
    img = Image.new("RGB", (1200, 630), COLORS["paper"])
    draw = ImageDraw.Draw(img)

    draw.rectangle((0, 0, 1200, 630), fill=COLORS["paper"])
    draw.rectangle((0, 0, 1200, 76), fill=COLORS["accent"])
    draw.text((64, 23), SCHOOL, font=font(24, True), fill="white")

    for x, y, w, h, color in [
        (760, 132, 310, 390, "#eef4ef"),
        (802, 98, 310, 390, "#ffffff"),
    ]:
        rounded_rectangle(draw, (x, y, x + w, y + h), 14, color, COLORS["line"], 2)
        draw.rectangle((x + 32, y + 40, x + w - 32, y + 48), fill=COLORS["accent"])
        for i in range(11):
            yy = y + 78 + i * 27
            draw.line((x + 32, yy, x + w - 32, yy), fill="#b8c9bf", width=2)
        draw.line((x + w // 2, y + 68, x + w // 2, y + h - 34), fill="#c5d0c8", width=2)
    draw.ellipse((1020, 410, 1138, 528), fill=COLORS["accent"], outline=COLORS["accent_dark"], width=4)
    draw.line((1052, 470, 1082, 500, 1110, 442), fill="white", width=12, joint="curve")

    draw.text((70, 154), TITLE, font=font(68, True), fill=COLORS["ink"])
    draw.text((74, 252), SUBTITLE, font=font(34, True), fill=COLORS["accent_dark"])

    bullets = ["B4 / A4 雙欄模板", "PDF 頁數檢查與預覽", "兩頁鎖定與格式校正"]
    for index, text in enumerate(bullets):
        y = 340 + index * 56
        draw.ellipse((78, y + 6, 104, y + 32), fill=COLORS["gold"])
        draw.text((122, y), text, font=font(30, True), fill=COLORS["ink"])

    rounded_rectangle(draw, (70, 530, 446, 584), 27, COLORS["accent"])
    draw.text((104, 544), "Word 考卷格式標準化工具", font=font(23, True), fill="white")
    draw.text((70, 596), "適合 LINE / Facebook 分享預覽", font=font(18), fill=COLORS["muted"])
    return img


def write_svg() -> None:
    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64">
  <rect width="64" height="64" rx="14" fill="{COLORS["accent"]}"/>
  <rect x="15" y="10" width="32" height="42" rx="4" fill="{COLORS["paper"]}" stroke="{COLORS["accent_dark"]}" stroke-width="2"/>
  <path d="M38 10v10h9" fill="#e8f0ea" stroke="{COLORS["accent_dark"]}" stroke-width="2"/>
  <path d="M21 25h20M21 32h20M21 39h14" stroke="{COLORS["accent"]}" stroke-width="3" stroke-linecap="round"/>
  <circle cx="43" cy="43" r="10" fill="{COLORS["accent"]}" stroke="{COLORS["accent_dark"]}" stroke-width="2"/>
  <path d="M38 43l4 4 7-9" fill="none" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
</svg>'''
    (ASSET_DIR / "favicon.svg").write_text(svg, encoding="utf-8")


def main() -> None:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    icon_16 = make_icon(16)
    icon_32 = make_icon(32)
    icon_48 = make_icon(48)
    icon_180 = make_icon(180)
    icon_192 = make_icon(192)
    icon_512 = make_icon(512)
    icon_192_mask = make_icon(192, maskable=True)
    icon_512_mask = make_icon(512, maskable=True)

    icon_16.save(ASSET_DIR / "favicon-16.png")
    icon_32.save(ASSET_DIR / "favicon-32.png")
    icon_48.save(ASSET_DIR / "favicon-48.png")
    icon_180.save(ASSET_DIR / "apple-touch-icon.png")
    icon_192.save(ASSET_DIR / "icon-192.png")
    icon_512.save(ASSET_DIR / "icon-512.png")
    icon_192_mask.save(ASSET_DIR / "icon-192-maskable.png")
    icon_512_mask.save(ASSET_DIR / "icon-512-maskable.png")
    icon_32.save(ASSET_DIR / "favicon.ico", sizes=[(16, 16), (32, 32), (48, 48)], append_images=[icon_16, icon_48])
    write_svg()
    make_og().save(ASSET_DIR / "og-exam-format.png", optimize=True)
    print(f"generated assets -> {ASSET_DIR}")


if __name__ == "__main__":
    main()
