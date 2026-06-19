from PIL import Image, ImageDraw, ImageFont
import os

OUT = r"C:\Users\VJ872NS\OneDrive - EY\Desktop\VIT_Report\landing_pages"
os.makedirs(OUT, exist_ok=True)
FD  = "C:/Windows/Fonts/"

# ── Palette ───────────────────────────────────────────────────────────────
NAVY1  = (13,  43,  85);  NAVY2  = (5,  18,  45)
GOLD   = (212, 160,  23); GOLD2  = (220, 185, 70)
WHITE  = (255, 255, 255); WDIM   = (170, 195, 220)
NAVY_C = (26,  65, 140);  GREEN_C = (15,  90,  65)
PURP_C = (80,  20, 140);  RED_C   = (155,  35,  35)

# ── Helpers ───────────────────────────────────────────────────────────────
def fnt(names, sz):
    for n in names:
        try: return ImageFont.truetype(FD + n, sz)
        except: pass
    return ImageFont.load_default()

SERIF = ['georgiab.ttf','georgia.ttf','timesbd.ttf','arialbd.ttf']
BOLD  = ['arialbd.ttf','calibrib.ttf','verdanab.ttf']
REG   = ['arial.ttf','calibri.ttf','verdana.ttf']

def vgrad(d, w, h, c1, c2):
    for y in range(h):
        t = y / max(h-1, 1)
        d.line([(0,y),(w-1,y)], fill=tuple(int(c1[i]+(c2[i]-c1[i])*t) for i in range(3)))

def rr(d, x1,y1,x2,y2, r, fill):
    d.rectangle([x1+r,y1,x2-r,y2], fill=fill)
    d.rectangle([x1,y1+r,x2,y2-r], fill=fill)
    for ex,ey in [(x1,y1),(x2-2*r,y1),(x1,y2-2*r),(x2-2*r,y2-2*r)]:
        d.ellipse([ex,ey,ex+2*r,ey+2*r], fill=fill)

def rings(d, cx, cy, start_r, count, step, base_color):
    for i in range(count):
        rs = start_r + i * step
        bc = base_color
        c  = (min(255,bc[0]+i*5), min(255,bc[1]+i*8), min(255,bc[2]+i*14))
        d.ellipse([cx-rs,cy-rs,cx+rs,cy+rs], outline=c, width=2)


# ══════════════════════════════════════════════════════════════════════════
# COVER  1280 × 720
# ══════════════════════════════════════════════════════════════════════════
W, H = 1280, 720
img = Image.new('RGB',(W,H))
d   = ImageDraw.Draw(img)
vgrad(d, W, H, NAVY1, NAVY2)

# Decorative rings – top-right corner
rings(d, W+20, -70, 160, 8, 62, (18,55,100))

# Fonts
f18 = fnt(BOLD, 18);  f22 = fnt(BOLD, 22)
f96 = fnt(SERIF,96);  f80 = fnt(SERIF,80)
f20 = fnt(BOLD, 20);  f16 = fnt(REG,  16); f14 = fnt(REG, 14); f13 = fnt(REG, 13)

# ── Left-side content ─────────────────────────────────────────────────────
# Gold accent bar
d.rectangle([80,62,370,67], fill=GOLD)
d.text((80,76), "VIT CHENNAI  *  ENGINEERING MATHEMATICS", font=f18, fill=GOLD)

# Main heading
d.text((80,125), "Complete",     font=f96, fill=WHITE)
d.text((80,235), "Study Bundle", font=f80, fill=GOLD)

# Tagline
d.text((80,350), "4 Guides   28 Modules   100+ Solved Problems   Print-Ready PDF",
       font=f22, fill=WDIM)

# Gold divider
d.rectangle([80,395,870,400], fill=GOLD)

# ── Right-side feature list ───────────────────────────────────────────────
feats = [
    "* Formula grids for every topic",
    "* Step-by-step worked examples",
    "* Colour-coded module structure",
    "* Master formula reference sheet",
    "* MathJax-rendered equations",
    "* Print as PDF in one click",
]
d.rectangle([930,108,934,358], fill=GOLD)   # vertical gold bar
for i, feat in enumerate(feats):
    y = 120 + i * 42
    d.text((952, y), feat, font=f16, fill=WDIM)

# ── Course cards (bottom) ─────────────────────────────────────────────────
courses = [
    ("BMAT101L", "Calculus &",     "Applications",  NAVY_C),
    ("BMAT102L", "DE &",           "Transforms",    GREEN_C),
    ("BMAT201L", "Complex Vars",   "& Lin. Alg.",   PURP_C),
    ("BMAT202L", "Probability",    "& Statistics",  RED_C),
]
CW,CH,CX0,CY0,CGAP = 265,190,80,455,20

for i,(code,l1,l2,col) in enumerate(courses):
    x1 = CX0 + i*(CW+CGAP);  y1 = CY0
    x2 = x1+CW;              y2 = y1+CH
    rr(d, x1,y1,x2,y2, 12, col)
    d.rectangle([x1+12,y1,x2-12,y1+5], fill=GOLD)          # gold top accent
    d.text((x1+14, y1+13), code, font=f20, fill=GOLD)       # code
    d.text((x1+14, y1+50), l1,   font=f16, fill=WHITE)      # name line 1
    d.text((x1+14, y1+70), l2,   font=f16, fill=WHITE)      # name line 2
    d.text((x1+14, y1+112),"7 Modules  *  Solved Examples", font=f13, fill=(160,190,215))
    d.text((x1+14, y1+130),"Formula Sheets  *  Print-Ready",font=f13, fill=(160,190,215))

# Bottom-right branding
d.text((W-300, H-28), "aarthir05.gumroad.com", font=f13, fill=(70,100,140))

cover_path = os.path.join(OUT, "cover.png")
img.save(cover_path, 'PNG', dpi=(72,72))
print(f"[OK] cover.png  1280x720  ->  {cover_path}")


# ══════════════════════════════════════════════════════════════════════════
# THUMBNAIL  600 × 600
# ══════════════════════════════════════════════════════════════════════════
TW,TH = 600,600
timg = Image.new('RGB',(TW,TH))
td   = ImageDraw.Draw(timg)
vgrad(td, TW, TH, NAVY1, NAVY2)

# Decorative rings – right edge
rings(td, TW+30, TH//2, 110, 6, 58, (18,55,100))

# Fonts
ft110 = fnt(SERIF,110); ft62 = fnt(SERIF,62)
ft22  = fnt(BOLD, 22);  ft16 = fnt(REG,  16); ft14 = fnt(REG, 14)

# Gold bar
td.rectangle([50,75,215,81], fill=GOLD)

# "VIT" – huge gold
td.text((50,85),  "VIT",         font=ft110, fill=GOLD)

# "Engineering" / "Mathematics"
td.text((50,213), "Engineering", font=ft62,  fill=WHITE)
td.text((50,290), "Mathematics", font=ft62,  fill=WHITE)

# Gold divider
td.rectangle([50,374,410,379], fill=GOLD)

# Sub-label
td.text((50,387), "Complete Study Bundle", font=ft22, fill=GOLD2)

# 4 course chips
chip_w,chip_h,chip_gap = 118,42,10
chip_y0 = 442
chip_x0 = 50
pairs = [(NAVY_C,"BMAT101L"),(GREEN_C,"BMAT102L"),(PURP_C,"BMAT201L"),(RED_C,"BMAT202L")]

for i,(col,code) in enumerate(pairs):
    cx1 = chip_x0 + i*(chip_w+chip_gap)
    rr(td, cx1,chip_y0,cx1+chip_w,chip_y0+chip_h, 8, col)
    td.rectangle([cx1+8,chip_y0,cx1+chip_w-8,chip_y0+4], fill=GOLD)
    td.text((cx1+8, chip_y0+11), code, font=ft14, fill=WHITE)

# Bottom tagline
td.text((50,508), "VIT Chennai  *  4 Guides  *  28 Modules  *  100+ Problems",
        font=ft14, fill=(140,170,200))

thumb_path = os.path.join(OUT, "thumbnail.png")
timg.save(thumb_path, 'PNG', dpi=(72,72))
print(f"[OK] thumbnail.png  600x600  ->  {thumb_path}")
