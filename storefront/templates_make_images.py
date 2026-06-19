"""
templates_make_images.py
Generates cover (1280x720) and thumbnail (600x600) for the
Aesthetic Slide Templates Gumroad product.
"""
from PIL import Image, ImageDraw, ImageFont
import os, math, random

OUT = r"C:\Users\VJ872NS\OneDrive - EY\Desktop\VIT_Report\storefront"
os.makedirs(OUT, exist_ok=True)
FD  = "C:/Windows/Fonts/"

# ── Palette ───────────────────────────────────────────────────────────────
BG1    = ( 18,   8,  38)   # deep purple-black
BG2    = (  8,   4,  22)   # near-black
PINK   = (240,  80, 160)
PURP   = (170,  80, 230)
CYAN   = ( 80, 210, 210)
GOLD   = (240, 200,  80)
CREAM  = (255, 250, 235)
MINT   = (160, 235, 195)
WDIM   = (210, 190, 240)

THEME_COLORS = [
    ((255,182,193), (216,191,216), "soft_bloom",   "Pastel Study"),
    ((238,228,210), (220,230,200), "raw_linen",    "Notion Minimal"),
    (( 15, 22, 65), ( 80,120,220), "dark_ether",   "Night Sky"),
    ((230,205,170), ( 90,125, 75), "wild_sage",    "Cottagecore"),
    (( 12,  5, 28), (255, 60,150), "neon_static",  "Y2K Retro"),
    ((255,220,230), (255,240,150), "sugar_bear",   "Sanrio Bear"),
]

# ── Helpers ───────────────────────────────────────────────────────────────
def fnt(names, sz):
    for n in names:
        try:    return ImageFont.truetype(FD + n, sz)
        except: pass
    return ImageFont.load_default()

SERIF = ['georgiab.ttf','georgia.ttf','timesbd.ttf','arialbd.ttf']
BOLD  = ['arialbd.ttf','calibrib.ttf','verdanab.ttf']
REG   = ['arial.ttf','calibri.ttf','verdana.ttf']

def vgrad(d, w, h, c1, c2):
    for y in range(h):
        t = y / max(h-1, 1)
        d.line([(0,y),(w-1,y)], fill=tuple(int(c1[i]+(c2[i]-c1[i])*t) for i in range(3)))

def circle(d, cx, cy, r, fill, outline=None, width=1):
    d.ellipse([cx-r, cy-r, cx+r, cy+r], fill=fill, outline=outline, width=width)

def rr(d, x1, y1, x2, y2, r, fill):
    d.rectangle([x1+r, y1, x2-r, y2], fill=fill)
    d.rectangle([x1, y1+r, x2, y2-r], fill=fill)
    for ex, ey in [(x1,y1),(x2-2*r,y1),(x1,y2-2*r),(x2-2*r,y2-2*r)]:
        d.ellipse([ex,ey,ex+2*r,ey+2*r], fill=fill)

def stars(d, w, h, count=80, seed=42):
    random.seed(seed)
    for _ in range(count):
        x, y = random.randint(0,w), random.randint(0,h)
        r    = random.uniform(1.2, 3.5)
        br   = random.randint(160, 255)
        circle(d, x, y, r, (br, br, min(255, br+50)))


# ══════════════════════════════════════════════════════════════════════════
# COVER  1280 × 720 px  72 DPI
# ══════════════════════════════════════════════════════════════════════════
W, H = 1280, 720
img = Image.new('RGB', (W, H))
d   = ImageDraw.Draw(img)

# Background gradient
vgrad(d, W, H, BG1, BG2)

# Star field
stars(d, W, H, 90)

# Decorative glowing rings top-right
for i in range(9):
    rs = 120 + i*70
    alpha = max(40, 140 - i*16)
    outline_col = (min(255,80+i*10), min(255,40+i*8), min(255,120+i*10))
    circle(d, W+40, -60, rs, None, outline_col, 2)

# ── Left side: text block ─────────────────────────────────────────────────
f_tag  = fnt(BOLD,  18)
f_h1   = fnt(SERIF, 88)
f_h2   = fnt(SERIF, 60)
f_sub  = fnt(BOLD,  22)
f_feat = fnt(REG,   18)
f_sm   = fnt(REG,   15)

# Tag
d.rectangle([72, 60, 72+300, 65], fill=PINK)
d.text((72, 76), "AESTHETIC  ✦  POWERPOINT  ✦  TEMPLATES", font=f_tag, fill=PINK)

# Heading
d.text((72, 124), "Six Cute",   font=f_h1, fill=CREAM)
d.text((72, 228), "Slide Packs",font=f_h2, fill=PURP)

# Subline
d.text((72, 328), "6 Themes  ·  10 Slides Each  ·  Instant .pptx Download",
       font=f_sub, fill=WDIM)

# Gold divider
d.rectangle([72, 372, 780, 377], fill=GOLD)

# Feature list
feats = [
    "✦  Pastel, Minimal, Night Sky, Cottagecore, Y2K, Sanrio",
    "✦  Fully editable — replace text and go",
    "✦  16:9 widescreen · System fonts · No plugins",
    "✦  Title, Agenda, Cards, Stats, Quote + more",
]
for i, feat in enumerate(feats):
    d.text((72, 392 + i*36), feat, font=f_feat, fill=WDIM)

# ── Right side: 6 mini theme chips ───────────────────────────────────────
chip_x = 830
chip_y_start = 120
chip_w, chip_h = 380, 72
gap = 16

for i, (bg, accent, name, label) in enumerate(THEME_COLORS):
    cy = chip_y_start + i * (chip_h + gap)
    # chip background
    rr(d, chip_x, cy, chip_x+chip_w, cy+chip_h, 10, bg)
    # accent bar on left
    rr(d, chip_x, cy, chip_x+8, cy+chip_h, 4, accent)
    # filename
    f_chip  = fnt(BOLD, 19)
    f_label = fnt(REG,  14)
    # text colour: light or dark depending on bg brightness
    lum = 0.299*bg[0] + 0.587*bg[1] + 0.114*bg[2]
    tc  = (30,20,40) if lum > 120 else CREAM
    d.text((chip_x+22, cy+10),    name,  font=f_chip,  fill=tc)
    d.text((chip_x+22, cy+38),    label, font=f_label, fill=accent)

# Bottom badge
badge_text = "All 6 files · Instant Download · No subscription"
d.rectangle([72, H-50, 72+580, H-16], fill=(255,255,255,0))
d.text((72, H-46), badge_text, font=f_sm, fill=(160,150,190))

# Save
img.save(os.path.join(OUT, "templates_cover.png"), dpi=(72,72))
print("[OK] templates_cover.png  (1280×720)")


# ══════════════════════════════════════════════════════════════════════════
# THUMBNAIL  600 × 600 px  72 DPI
# ══════════════════════════════════════════════════════════════════════════
S = 600
img2 = Image.new('RGB', (S, S))
d2   = ImageDraw.Draw(img2)

# BG
vgrad(d2, S, S, BG1, BG2)
stars(d2, S, S, 50, seed=7)

# Decorative ring
for i in range(6):
    rs = 250 + i*35
    circle(d2, S+20, -20, rs, None, (80+i*8, 40+i*6, 120+i*10), 2)

# 6 small coloured dots — one per theme
dot_r = 18
angles = [i * (360/6) for i in range(6)]
cx, cy = 300, 220
orbit  = 130
for i, (bg, accent, _, _) in enumerate(THEME_COLORS):
    ang = math.radians(angles[i] - 90)
    dx  = int(cx + orbit * math.cos(ang))
    dy  = int(cy + orbit * math.sin(ang))
    circle(d2, dx, dy, dot_r+4, (30,15,60))
    circle(d2, dx, dy, dot_r,   bg)
    circle(d2, dx, dy, dot_r-6, accent)

# Centre icon circle
circle(d2, cx, cy, 52, (40,20,80))
circle(d2, cx, cy, 46, PURP)
f_icon = fnt(BOLD, 32)
d2.text((cx-10, cy-20), "✦", font=f_icon, fill=GOLD)

# Main text
f_t1 = fnt(SERIF, 48)
f_t2 = fnt(BOLD,  20)
f_t3 = fnt(REG,   15)

# "slide" top, bold pink
d2.text((0,0), "", font=f_t1, fill=PINK)   # dummy call to warm up
w1 = d2.textlength("slide", font=f_t1)
d2.text(((S-w1)//2, 370), "slide", font=f_t1, fill=PINK)

w2 = d2.textlength("aesthetic", font=f_t2)
d2.text(((S-w2)//2, 432), "aesthetic", font=f_t2, fill=PURP)

w3 = d2.textlength("6 themes · 10 slides each", font=f_t3)
d2.text(((S-w3)//2, 468), "6 themes · 10 slides each", font=f_t3, fill=WDIM)

# Bottom accent line
d2.rectangle([(S//2-100, S-38), (S//2+100, S-34)], fill=GOLD)

# Save
img2.save(os.path.join(OUT, "templates_thumbnail.png"), dpi=(72,72))
print("[OK] templates_thumbnail.png  (600×600)")
print(f"\nSaved to: {OUT}")
