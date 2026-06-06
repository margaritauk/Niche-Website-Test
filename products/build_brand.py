#!/usr/bin/env python3
"""
Generate the SHOP's brand assets (logo, icons, banners, covers, bundle hero).

Brand name + tagline are variables below — change them and re-run to rebrand everything
in one step. Output -> shop/assets/

Usage: python3 build_brand.py
"""
import os
from PIL import Image, ImageDraw, ImageFont

# ============================ BRAND (edit here to rebrand) ============================
BRAND   = "LedgerLite"
BRAND_2 = ("Ledger", "Lite")          # two-tone split for the wordmark
TAGLINE = "Bookkeeping that does your taxes for you"
SUBTAG  = "Spreadsheet trackers for hosts, sellers & the self-employed"
# =====================================================================================

NAVY=(31,58,95); TEAL=(44,122,123); TEAL_L=(72,160,162); LIGHT=(235,242,247)
LIGHTER=(245,249,252); WHITE=(255,255,255); INK=(26,32,44); GREY=(113,128,150)
FROOT="/usr/share/fonts/truetype/dejavu"
def F(sz, bold=True): return ImageFont.truetype(f"{FROOT}/DejaVuSans{'-Bold' if bold else ''}.ttf", sz)

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "shop", "assets")
os.makedirs(OUT, exist_ok=True)

def tlen(d, t, f): return d.textlength(t, font=f)

def pill(d, x, y, label, font, fg=WHITE, bg=TEAL, padx=30, pady=16):
    w = tlen(d, label, font); h = font.getbbox("Ay")[3]
    d.rounded_rectangle([x, y, x+w+2*padx, y+h+2*pady], radius=(h+2*pady)//2, fill=bg)
    d.text((x+padx, y+pady-2), label, font=font, fill=fg)
    return x+w+2*padx

def mark(d, x, y, s, on_dark=True):
    """The LedgerLite mark: rounded square with three rising bars + a check badge."""
    bg = WHITE if on_dark else NAVY
    d.rounded_rectangle([x, y, x+s, y+s], radius=int(s*0.22), fill=bg)
    # three rising bars
    pad = s*0.20; bw = s*0.15; gap = s*0.075
    base = y+s-pad
    heights = [0.26, 0.42, 0.58]
    cols = [TEAL_L, TEAL, NAVY if on_dark else TEAL]
    bx = x+pad
    for h, c in zip(heights, cols):
        d.rounded_rectangle([bx, base-s*h, bx+bw, base], radius=int(bw*0.25),
                            fill=(c if on_dark else WHITE if c==TEAL else c))
        bx += bw+gap
    # check badge top-right
    r = s*0.18; cx, cy = x+s-pad*0.5, y+pad*0.6
    d.ellipse([cx-r, cy-r, cx+r, cy+r], fill=TEAL)
    d.line([cx-r*0.45, cy, cx-r*0.05, cy+r*0.4], fill=WHITE, width=max(3,int(s*0.03)))
    d.line([cx-r*0.05, cy+r*0.4, cx+r*0.55, cy-r*0.45], fill=WHITE, width=max(3,int(s*0.03)))

def icon(size, on_dark=True, fname=None):
    img = Image.new("RGB", (size, size), NAVY if on_dark else WHITE)
    d = ImageDraw.Draw(img)
    m = int(size*0.16)
    mark(d, m, m, size-2*m, on_dark=on_dark)
    if fname: img.save(os.path.join(OUT, fname)); print("wrote", fname)
    return img

def wordmark(w, h, on_dark=True, tagline=True, fname=None):
    img = Image.new("RGB", (w, h), NAVY if on_dark else WHITE)
    d = ImageDraw.Draw(img)
    s = int(h*0.52)
    my = (h - s)//2 if not tagline else int(h*0.20)
    mx = int(w*0.06)
    mark(d, mx, my, s, on_dark=on_dark)
    # wordmark text
    tx = mx + s + int(w*0.03)
    fname_f = F(int(h*0.30))
    a, b = BRAND_2
    c1 = WHITE if on_dark else NAVY
    ty = my + (s - F(int(h*0.30)).getbbox("Ay")[3])//2 if not tagline else my - int(h*0.02)
    d.text((tx, ty), a, font=fname_f, fill=c1)
    wa = tlen(d, a, fname_f)
    d.text((tx+wa, ty), b, font=fname_f, fill=TEAL_L if on_dark else TEAL)
    if tagline:
        tf = F(int(h*0.105), False)
        d.text((tx, my + s - int(h*0.04)), TAGLINE, font=tf, fill=(180,205,224) if on_dark else GREY)
    if fname: img.save(os.path.join(OUT, fname)); print("wrote", fname)
    return img

def banner(w, h, fname, mini=False):
    img = Image.new("RGB", (w, h), NAVY)
    d = ImageDraw.Draw(img)
    s = int(h*0.46)
    mx = int(w*0.04); my = (h - s)//2
    mark(d, mx, my, s, on_dark=True)
    tx = mx + s + int(w*0.025)
    nf = F(int(h*0.26))
    a, b = BRAND_2
    ty = my + int(h*0.06)
    d.text((tx, ty), a, font=nf, fill=WHITE); wa = tlen(d, a, nf)
    d.text((tx+wa, ty), b, font=nf, fill=TEAL_L)
    tf = F(int(h*0.11), False)
    d.text((tx, ty + int(h*0.30)), TAGLINE, font=tf, fill=(180,205,224))
    if not mini:
        # right-side badges
        bf = F(int(h*0.075))
        labels = ["Excel + Google Sheets", "One-time purchase", "Instant download"]
        # stack badges on the right
        bx0 = int(w*0.66); by = my + int(h*0.10)
        for lab in labels:
            wlab = tlen(d, lab, bf)
            pill(d, w-int(w*0.03)-wlab-60, by, lab, bf)
            by += int(h*0.20)
    img.save(os.path.join(OUT, fname)); print("wrote", fname)

def gumroad_cover(w, h, fname):
    img = Image.new("RGB", (w, h), NAVY)
    d = ImageDraw.Draw(img)
    s = int(h*0.30)
    mark(d, int(w*0.06), int(h*0.16), s, on_dark=True)
    nf = F(int(h*0.13)); a, b = BRAND_2
    tx = int(w*0.06)+s+int(w*0.03); ty = int(h*0.18)
    d.text((tx, ty), a, font=nf, fill=WHITE); wa = tlen(d, a, nf)
    d.text((tx+wa, ty), b, font=nf, fill=TEAL_L)
    d.text((tx, ty+int(h*0.16)), TAGLINE, font=F(int(h*0.06), False), fill=(180,205,224))
    d.text((int(w*0.06), int(h*0.62)), SUBTAG, font=F(int(h*0.058), False), fill=WHITE)
    x = int(w*0.06)
    for lab in ["No subscription", "Excel & Google Sheets", "Set up in 10 min"]:
        x = pill(d, x, int(h*0.74), lab, F(int(h*0.045))) + 24
    img.save(os.path.join(OUT, fname)); print("wrote", fname)

def bundle_hero(fname):
    S=2000; img=Image.new("RGB",(S,S),NAVY); d=ImageDraw.Draw(img)
    mark(d, 120, 130, 240, on_dark=True)
    d.text((400, 150), BRAND, font=F(70), fill=WHITE)
    d.text((400, 250), "Everything Bundle", font=F(96), fill=TEAL_L)
    d.text((120, 470), "All 10 bookkeeping & tax trackers", font=F(66), fill=WHITE)
    trades = ["Airbnb / rentals","Landlords","Etsy & Amazon sellers","Freelancers / 1099",
              "Real estate agents","Photographers","House cleaners","Hair & beauty pros",
              "Rideshare & delivery","Food trucks & vendors"]
    y=600
    for t in trades:
        d.ellipse([130,y+8,178,y+56],fill=TEAL); d.line([143,y+33,158,y+48],fill=WHITE,width=7)
        d.line([158,y+48,176,y+18],fill=WHITE,width=7)
        d.text((210,y), t, font=F(46,False), fill=WHITE); y+=108
    d.rounded_rectangle([120,1720,S-120,1880],radius=24,fill=TEAL)
    d.text((160,1755),"All ten — one purchase",font=F(56),fill=WHITE)
    pf=F(60); lab="$69"; d.text((S-160-tlen(d,lab,pf),1748),lab,font=pf,fill=WHITE)
    img.save(os.path.join(OUT, fname)); print("wrote", fname)

if __name__ == "__main__":
    icon(500,  on_dark=True,  fname="icon-500.png")          # Etsy shop icon / Gumroad avatar
    icon(1000, on_dark=True,  fname="icon-1000.png")
    icon(500,  on_dark=False, fname="icon-500-light.png")
    wordmark(1600, 500, on_dark=False, fname="logo-on-white.png")
    wordmark(1600, 500, on_dark=True,  fname="logo-on-navy.png")
    banner(3360, 840,  "etsy-cover-3360x840.png")            # Etsy big cover photo
    banner(1200, 300,  "etsy-banner-mini-1200x300.png", mini=True)
    gumroad_cover(1280, 720, "gumroad-cover-1280x720.png")
    bundle_hero("everything-bundle-hero-2000.png")
    print("\nBrand:", BRAND, "—", TAGLINE)
