#!/usr/bin/env python3
"""
Generate ready-to-upload Etsy/Gumroad product images (5 PNGs per product, 2000x2000).

Pulls the product name, entity, feature data and the live sample numbers from the same
VARIANTS config as build_tracker.py, so the images always match the product. Output goes to
each product's  images/  folder:  01-hero.png ... 05-summary.png

Usage:
  python3 build_mockups.py            # all products
  python3 build_mockups.py str driver # specific variants
"""
import os, sys
from PIL import Image, ImageDraw, ImageFont
from build_tracker import VARIANTS

# ---- palette
NAVY=(31,58,95); TEAL=(44,122,123); LIGHT=(235,242,247); LIGHTER=(245,249,252)
WHITE=(255,255,255); INK=(26,32,44); GREY=(113,128,150); GREEN=(47,133,90); RED=(197,48,48)
FROOT="/usr/share/fonts/truetype/dejavu"
def F(size, bold=True):
    return ImageFont.truetype(f"{FROOT}/DejaVuSans{'-Bold' if bold else ''}.ttf", size)

S = 2000  # canvas size

# ---- per-product display copy (name + hero subline + image-2 hook)
DISPLAY = {
    "str":         ("Short-Term Rental Tracker",      "For Airbnb & VRBO hosts",                 "Profit by property"),
    "landlord":    ("Rental Property Tracker",        "For long-term landlords",                 "Profit by property"),
    "seller":      ("Online Seller Bookkeeping",      "For Etsy & Amazon sellers",               "Real profit, after fees"),
    "freelancer":  ("Freelancer Bookkeeping",         "For freelancers & 1099 contractors",      "Profit by client"),
    "realtor":     ("Real Estate Agent Tracker",      "For realtors & brokers",                  "Take-home, after your split"),
    "photographer":("Photographer Bookkeeping",       "For photographers & videographers",       "Profit by shoot"),
    "cleaning":    ("Cleaning Business Tracker",       "For house cleaners & companies",          "Profit by client"),
    "beauty":      ("Hair Stylist & Beauty Tracker",  "For stylists, barbers & nail techs",      "Service + retail + tips"),
    "driver":      ("Rideshare & Delivery Tracker",   "For Uber, Lyft & DoorDash drivers",       "Which app pays best"),
    "foodtruck":   ("Food Truck & Vendor Tracker",    "For food trucks & market vendors",        "Profit by event"),
}

# ---------------------------------------------------------------- helpers
def draw():
    img = Image.new("RGB", (S, S), WHITE)
    return img, ImageDraw.Draw(img)

def wrap(d, text, font, max_w):
    words, lines, cur = text.split(), [], ""
    for w in words:
        t = (cur + " " + w).strip()
        if d.textlength(t, font=font) <= max_w:
            cur = t
        else:
            if cur: lines.append(cur)
            cur = w
    if cur: lines.append(cur)
    return lines

def text_block(d, x, y, text, font, fill, max_w, leading=1.12, align="left", center_x=None):
    lines = wrap(d, text, font, max_w)
    h = font.getbbox("Ay")[3]
    for ln in lines:
        if align == "center" and center_x is not None:
            w = d.textlength(ln, font=font); d.text((center_x - w/2, y), ln, font=font, fill=fill)
        else:
            d.text((x, y), ln, font=font, fill=fill)
        y += int(h * leading)
    return y

def pill(d, x, y, label, font, fg=WHITE, bg=TEAL, padx=34, pady=18):
    w = d.textlength(label, font=font); h = font.getbbox("Ay")[3]
    d.rounded_rectangle([x, y, x + w + 2*padx, y + h + 2*pady], radius=(h+2*pady)//2, fill=bg)
    d.text((x + padx, y + pady - 2), label, font=font, fill=fg)
    return x + w + 2*padx

def money(v):
    s = f"${abs(v):,.0f}" if abs(v) >= 1000 or v == int(v) else f"${abs(v):,.2f}"
    return f"(${abs(v):,.0f})" if v < 0 else s

# ---- compute live numbers from samples
def numbers(cfg):
    s = cfg["samples"]
    inc = round(sum(r[5] for r in s if r[2] == "Income"), 2)
    exp = round(sum(r[5] for r in s if r[2] == "Expense"), 2)
    rows = []
    for e in dict.fromkeys(r[1] for r in s):
        i = round(sum(r[5] for r in s if r[1] == e and r[2] == "Income"), 2)
        x = round(sum(r[5] for r in s if r[1] == e and r[2] == "Expense"), 2)
        rows.append((e, i, x, round(i - x, 2)))
    return inc, exp, round(inc - exp, 2), rows

# ---------------------------------------------------------------- the 5 images
def dashboard_card(d, x, y, w, net, inc, exp, big=True):
    """A rounded 'software' card showing NET profit + income/expense chips."""
    h = 470 if big else 380
    d.rounded_rectangle([x, y, x + w, y + h], radius=36, fill=LIGHTER, outline=(203,213,224), width=3)
    d.text((x + 50, y + 40), "NET PROFIT", font=F(40), fill=TEAL)
    d.text((x + 50, y + 95), money(net), font=F(150), fill=TEAL)
    cy = y + 300
    chip_w = (w - 150) // 2
    for i, (lab, val, col) in enumerate([("Income", inc, GREEN), ("Expenses", exp, RED)]):
        cx = x + 50 + i * (chip_w + 50)
        d.rounded_rectangle([cx, cy, cx + chip_w, cy + 110], radius=20, fill=WHITE, outline=(203,213,224), width=2)
        d.text((cx + 30, cy + 18), lab, font=F(34, False), fill=GREY)
        d.text((cx + 30, cy + 55), money(val), font=F(46), fill=col)

def img_hero(cfg, name, sub):
    img, d = draw()
    d.rectangle([0, 0, S, 1080], fill=NAVY)
    text_block(d, 120, 150, name, F(112), WHITE, S - 240, leading=1.05)
    d.text((120, 470), sub, font=F(52, False), fill=(180,205,224))
    # badges
    x = 120
    for b in ["Excel + Google Sheets", "One-time purchase", "Instant download"]:
        x = pill(d, x, 600, b, F(36)) + 28
    d.text((120, 770), "Bookkeeping that does your taxes for you.", font=F(46), fill=WHITE)
    inc, exp, net, _ = numbers(cfg)
    dashboard_card(d, 120, 1200, S - 240, net, inc, exp)
    d.text((120, 1740), "IRS " + cfg.get("form", "Schedule E") + "  ·  tracks up to 5  ·  set up in 10 minutes",
           font=F(40, False), fill=GREY)
    return img

def img_profit(cfg, name, hook):
    img, d = draw()
    inc, exp, net, rows = numbers(cfg)
    text_block(d, 120, 130, "See your real profit", F(96), NAVY, S - 240, leading=1.05)
    d.text((120, 360), hook, font=F(54, False), fill=TEAL)
    # table
    x0, y0, w = 120, 520, S - 240
    cols = [0.40, 0.20, 0.20, 0.20]
    cx = [x0 + sum(cols[:i]) * w for i in range(4)]
    ent_label = cfg.get("entity", "Property")
    d.rounded_rectangle([x0, y0, x0 + w, y0 + 90], radius=12, fill=TEAL)
    for c, lab in zip(cx, [ent_label, "Income", "Expenses", "Net"]):
        d.text((c + 28, y0 + 24), lab, font=F(40), fill=WHITE)
    y = y0 + 90
    shown = rows[:4]
    for i, (e, ii, xx, nn) in enumerate(shown):
        d.rectangle([x0, y, x0 + w, y + 120], fill=LIGHTER if i % 2 == 0 else WHITE)
        nm = (e[:22] + "…") if len(e) > 23 else e
        d.text((cx[0] + 28, y + 36), nm, font=F(38, False), fill=INK)
        d.text((cx[1] + 28, y + 36), money(ii), font=F(38, False), fill=GREEN)
        d.text((cx[2] + 28, y + 36), money(xx), font=F(38, False), fill=RED)
        d.text((cx[3] + 28, y + 36), money(nn), font=F(40), fill=(TEAL if nn >= 0 else RED))
        y += 120
    d.rectangle([x0, y, x0 + w, y + 130], fill=LIGHT)
    d.text((cx[0] + 28, y + 40), "TOTAL", font=F(42), fill=NAVY)
    d.text((cx[1] + 28, y + 40), money(inc), font=F(40), fill=GREEN)
    d.text((cx[2] + 28, y + 40), money(exp), font=F(40), fill=RED)
    d.text((cx[3] + 28, y + 40), money(net), font=F(44), fill=TEAL)
    d.text((120, y + 200), "↑ Fills in automatically — and maps to your IRS " + cfg.get("form","Schedule E") + ".",
           font=F(44, False), fill=GREY)
    d.text((120, y + 270), "Hand the summary straight to your accountant.", font=F(44, False), fill=GREY)
    return img

def img_steps(cfg, name):
    img, d = draw()
    d.rectangle([0, 0, S, S], fill=NAVY)
    text_block(d, 120, 150, "No accounting degree required", F(92), WHITE, S - 240, leading=1.05)
    steps = [("1", "Log it once", "Type in the amount and pick from a dropdown."),
             ("2", "It auto-categorizes", "Every entry lands on the right tax line."),
             ("3", "Reports fill in", "Summary + dashboard update instantly.")]
    y = 560
    for n, t, sub in steps:
        d.rounded_rectangle([120, y, S - 120, y + 360], radius=28, fill=(40,70,110))
        d.ellipse([170, y + 110, 310, y + 250], fill=TEAL)
        w = d.textlength(n, font=F(90)); d.text((240 - w/2, y + 130), n, font=F(90), fill=WHITE)
        d.text((380, y + 110), t, font=F(60), fill=WHITE)
        d.text((380, y + 205), sub, font=F(42, False), fill=(180,205,224))
        y += 420
    return img

def img_included(cfg, name):
    img, d = draw()
    text_block(d, 120, 150, "What you get", F(100), NAVY, S - 240)
    # file icons
    for i, (lab, col) in enumerate([("XLSX", GREEN), ("PDF", RED)]):
        x = 160 + i * 520
        d.rounded_rectangle([x, 420, x + 420, 820], radius=30, fill=LIGHTER, outline=(203,213,224), width=3)
        d.rounded_rectangle([x + 120, 470, x + 300, 560], radius=12, fill=col)
        w = d.textlength(lab, font=F(46)); d.text((x + 210 - w/2, 488), lab, font=F(46), fill=WHITE)
        sub = "The tracker" if lab == "XLSX" else "Setup guide"
        ww = d.textlength(sub, font=F(40, False)); d.text((x + 210 - ww/2, 620), sub, font=F(40, False), fill=INK)
    bullets = [
        "Automated 6-tab tracker (.xlsx)",
        f"Tracks up to 5 {cfg.get('entity_plural_lower','properties')}",
        "Step-by-step setup guide (PDF)",
        "Works in Excel, Google Sheets & Numbers",
        "Instant download — yours forever",
    ]
    y = 980
    for b in bullets:
        d.ellipse([130, y + 6, 184, y + 60], fill=TEAL)
        d.line([145, y + 33, 160, y + 48], fill=WHITE, width=8); d.line([160, y + 48, 178, y + 18], fill=WHITE, width=8)
        d.text((220, y), b, font=F(48, False), fill=INK)
        y += 110
    return img

def img_summary(cfg, name):
    img, d = draw()
    d.rectangle([0, 0, S, S], fill=NAVY)
    text_block(d, 120, 470, "No subscriptions. No logins. No monthly fees.", F(86), WHITE, S - 240, leading=1.12)
    d.text((120, 980), "Works in  Excel · Google Sheets · Apple Numbers", font=F(52, False), fill=(180,205,224))
    x = 120
    for b in ["One-time purchase", "Instant download", "Set up in 10 min"]:
        x = pill(d, x, 1120, b, F(38)) + 28
    d.text((120, 1380), name, font=F(56), fill=WHITE)
    d.text((120, 1470), "A deductible business tool — built for your trade.", font=F(44, False), fill=GREY)
    return img

BUILDERS = [("01-hero", img_hero, True), ("02-profit", img_profit, True),
            ("03-steps", img_steps, False), ("04-included", img_included, False),
            ("05-summary", img_summary, False)]

def build_images(cfg):
    name, sub, hook = DISPLAY[cfg["_key"]]
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), cfg["dir"], "images")
    os.makedirs(out, exist_ok=True)
    for fname, fn, needs_hook in BUILDERS:
        img = fn(cfg, name, hook if fn is img_profit else sub) if fn in (img_hero, img_profit) else fn(cfg, name)
        img.save(os.path.join(out, fname + ".png"))
    print("wrote 5 images ->", os.path.relpath(out))

if __name__ == "__main__":
    keys = sys.argv[1:] or list(VARIANTS)
    for k in keys:
        cfg = dict(VARIANTS[k]); cfg["_key"] = k
        build_images(cfg)
