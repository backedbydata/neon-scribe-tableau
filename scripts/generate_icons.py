"""
Generate ICO files for Neon Scribe (Tableau and Power BI versions).

Design:
- Dark navy background with neon cyan glowing pen/stylus (diagonal, top-right to bottom-left)
- Tableau version: Tableau's iconic dot-cross arrangement in orange (#E97627) bottom-right
- Power BI version: Power BI's stacked bar chart shape in yellow (#F2C811) bottom-right
- Both icons share the same pen design — only the brand mark differs
"""

from PIL import Image, ImageDraw, ImageFilter, ImageFont
import math
import os


def lerp(a, b, t):
    return a + (b - a) * t


def draw_base(size):
    """Dark navy background with subtle neon ambient glow."""
    s = size
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))

    # Base background
    bg = Image.new("RGBA", (s, s), (10, 10, 28, 255))
    draw = ImageDraw.Draw(bg)

    # Rounded rect to soften corners slightly at small sizes
    r = s // 8
    draw.rounded_rectangle([0, 0, s - 1, s - 1], radius=r, fill=(10, 10, 28, 255))

    return bg


def draw_pen(img, size, accent=(0, 230, 255), glow=(0, 120, 200)):
    """Draw a glowing neon pen/stylus diagonally across the icon."""
    s = size
    draw = ImageDraw.Draw(img)

    # Pen endpoints — tip at bottom-left, cap at top-right
    tip_x = int(s * 0.20)
    tip_y = int(s * 0.82)
    cap_x = int(s * 0.72)
    cap_y = int(s * 0.20)

    pw = max(2, int(s * 0.07))   # pen body half-width

    def thick_line(p1, p2, width, color):
        dx = p2[0] - p1[0]
        dy = p2[1] - p1[1]
        length = math.sqrt(dx * dx + dy * dy)
        if length == 0:
            return
        nx = -dy / length * width / 2
        ny = dx / length * width / 2
        pts = [
            (p1[0] + nx, p1[1] + ny),
            (p1[0] - nx, p1[1] - ny),
            (p2[0] - nx, p2[1] - ny),
            (p2[0] + nx, p2[1] + ny),
        ]
        draw.polygon(pts, fill=color)

    # Outer glow passes
    for extra, alpha in [(10, 25), (7, 45), (5, 70), (3, 100)]:
        thick_line((tip_x, tip_y), (cap_x, cap_y), pw + extra, (*glow, alpha))

    # Body
    thick_line((tip_x, tip_y), (cap_x, cap_y), pw, (*accent, 255))

    # Center glint (bright white highlight)
    thick_line((tip_x, tip_y), (cap_x, cap_y), max(1, pw // 3), (255, 255, 255, 200))

    # Pen nib / tip triangle
    angle = math.atan2(cap_y - tip_y, cap_x - tip_x)
    perp = angle + math.pi / 2
    nib_len = pw * 2.2
    nib_tip_x = tip_x - math.cos(angle) * nib_len
    nib_tip_y = tip_y - math.sin(angle) * nib_len
    side = pw * 0.6
    nib_pts = [
        (tip_x + math.cos(perp) * side, tip_y + math.sin(perp) * side),
        (tip_x - math.cos(perp) * side, tip_y - math.sin(perp) * side),
        (nib_tip_x, nib_tip_y),
    ]
    draw.polygon(nib_pts, fill=(*accent, 255))

    # Cap end — rounded bright dot
    cap_r = pw // 2 + max(1, s // 40)
    # glow
    for gr in [cap_r + 5, cap_r + 3, cap_r + 1]:
        draw.ellipse([cap_x - gr, cap_y - gr, cap_x + gr, cap_y + gr],
                     fill=(*accent, 60))
    draw.ellipse([cap_x - cap_r, cap_y - cap_r, cap_x + cap_r, cap_y + cap_r],
                 fill=(*accent, 255))

    # Small spark dots near nib tip
    for offset, r_frac, a in [
        (-0.05, 0.015, 200),
        (-0.08, 0.010, 150),
        (-0.03, 0.008, 120),
    ]:
        sx = int(nib_tip_x + math.cos(perp) * s * offset * 2 - math.cos(angle) * s * 0.02)
        sy = int(nib_tip_y + math.sin(perp) * s * offset * 2 - math.sin(angle) * s * 0.02)
        sr = max(1, int(s * r_frac))
        draw.ellipse([sx - sr, sy - sr, sx + sr, sy + sr], fill=(*accent, a))

    return img


def draw_tableau_mark(img, size, color=(233, 118, 39)):
    """
    Tableau's iconic mark: a cross arrangement of filled squares/dots.
    Central dot + 4 arms (top, bottom, left, right) + 4 diagonals = 9 dots total.
    Rendered at bottom-right corner of the icon.
    """
    s = size
    draw = ImageDraw.Draw(img)

    # Centre of the mark, bottom-right quadrant
    cx = int(s * 0.74)
    cy = int(s * 0.74)

    # Dot radius and spacing scaled to icon size
    dot_r = max(1, int(s * 0.045))
    spacing = int(s * 0.115)

    positions = [
        (0, 0),           # centre
        (0, -spacing),    # top
        (0,  spacing),    # bottom
        (-spacing, 0),    # left
        ( spacing, 0),    # right
        # shorter diagonals (Tableau's mark has them slightly shorter)
        (-spacing * 0.65, -spacing * 0.65),
        ( spacing * 0.65, -spacing * 0.65),
        (-spacing * 0.65,  spacing * 0.65),
        ( spacing * 0.65,  spacing * 0.65),
    ]
    # Diagonal dots are slightly smaller
    dot_sizes = [dot_r] * 5 + [max(1, int(dot_r * 0.75))] * 4

    # Glow pass
    glow_col = color
    for (dx, dy), dr in zip(positions, dot_sizes):
        px, py = cx + dx, cy + dy
        gr = dr + max(2, s // 32)
        draw.ellipse([px - gr, py - gr, px + gr, py + gr], fill=(*glow_col, 60))

    # Dots
    for (dx, dy), dr in zip(positions, dot_sizes):
        px, py = cx + dx, cy + dy
        draw.ellipse([px - dr, py - dr, px + dr, py + dr], fill=(*color, 235))
        # centre glint
        hr = max(1, dr // 3)
        draw.ellipse([px - hr, py - hr, px + hr, py + hr], fill=(255, 255, 255, 150))

    return img


def draw_powerbi_mark(img, size, color=(242, 200, 17)):
    """
    Power BI's iconic stacked bar / staircase chart.
    Three bars of increasing height, left to right.
    Rendered at bottom-right corner of the icon.
    """
    s = size
    draw = ImageDraw.Draw(img)

    # Anchor bottom-right
    mark_w = int(s * 0.28)
    mark_h = int(s * 0.28)
    right  = int(s * 0.88)
    bottom = int(s * 0.88)
    left   = right - mark_w
    top    = bottom - mark_h

    bar_count = 3
    gap = max(1, int(mark_w * 0.08))
    bar_w = (mark_w - gap * (bar_count - 1)) // bar_count

    # Heights: 40%, 70%, 100% of mark_h
    heights = [int(mark_h * 0.40), int(mark_h * 0.70), mark_h]

    for i in range(bar_count):
        bx0 = left + i * (bar_w + gap)
        bx1 = bx0 + bar_w
        bh  = heights[i]
        by0 = bottom - bh
        by1 = bottom

        # Glow
        gr = max(2, s // 40)
        draw.rectangle([bx0 - gr, by0 - gr, bx1 + gr, by1 + gr], fill=(*color, 55))
        # Bar fill
        draw.rectangle([bx0, by0, bx1, by1], fill=(*color, 235))
        # Top glint
        gh = max(1, bh // 5)
        draw.rectangle([bx0, by0, bx1, by0 + gh], fill=(255, 255, 255, 120))

    return img


def make_icon(size, brand):
    img = draw_base(size)
    img = draw_pen(img, size)

    if brand == "tableau":
        img = draw_tableau_mark(img, size, color=(233, 118, 39))
    else:
        img = draw_powerbi_mark(img, size, color=(242, 200, 17))

    # Soft glow bloom — blend a blurred copy under the sharp image
    blurred = img.filter(ImageFilter.GaussianBlur(radius=max(1, size // 28)))
    combined = Image.alpha_composite(blurred, img)
    return combined


def make_ico(brand, output_path):
    sizes = [16, 32, 48, 64, 128, 256]
    frames = [make_icon(sz, brand) for sz in sizes]

    frames[-1].save(
        output_path,
        format="ICO",
        sizes=[(f.width, f.height) for f in frames],
        append_images=frames[:-1],
    )
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    out_dir = os.path.join(os.path.dirname(__file__), "..", "assets")
    os.makedirs(out_dir, exist_ok=True)

    make_ico("tableau", os.path.join(out_dir, "neon-scribe-tableau.ico"))
    make_ico("powerbi",  os.path.join(out_dir, "neon-scribe-powerbi.ico"))

    # Export preview PNGs
    for brand, fname in [("tableau", "neon-scribe-tableau.ico"), ("powerbi", "neon-scribe-powerbi.ico")]:
        ico = Image.open(os.path.join(out_dir, fname))
        ico.save(os.path.join(out_dir, f"preview-{brand}.png"))
        print(f"Preview: preview-{brand}.png")

    print("Done.")
