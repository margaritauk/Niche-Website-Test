#!/usr/bin/env python3
"""
Generate vertical Pinterest/Instagram pins (1000x1500) for every product + bundle + the shop.
Pinterest is the top traffic source for spreadsheet/printable shops, so these are the
marketing engine. Output -> social/pins/

Usage: python3 build_pins.py
"""
import os
from PIL import Image, ImageDraw, ImageFont
from build_tracker import VARIANTS
from build_mockups import DISPLAY, numbers, money
from build_bundles import BUNDLES

NAVY=(31,58,95); TEAL=(44,122,123); TEAL_L=(72,160,162); WHITE=(255,255,255)
LIGHT=(235,242,247); LIGHTER=(245,249,252); INK=(26,32,44); GREY=(113,128,150)
GREEN=(47,133,90); RED=(197,48,48)
FROOT="/usr/share/fonts/truetype/dejavu"
def F(s,b=True): return ImageFont.truetype(f"{FROOT}/DejaVuSans{'-Bold' if b else ''}.ttf", s)
W,H=1000,1500
OUT=os.path.join(os.path.dirname(os.path.abspath(__file__)),"social","pins"); os.makedirs(OUT,exist_ok=True)

HOOKS={
 "str":"The Airbnb spreadsheet that does your taxes for you",
 "landlord":"Landlord bookkeeping, finally made simple",
 "seller":"Etsy seller? See your REAL profit after fees",
 "freelancer":"Freelancer bookkeeping that shows profit by client",
 "realtor":"Realtor commissions & taxes, handled",
 "photographer":"Photographers: see which shoots actually pay",
 "cleaning":"Cleaning business bookkeeping made easy",
 "beauty":"Hair & beauty pros: track services, retail & tips",
 "driver":"Uber & DoorDash drivers: which app pays best?",
 "foodtruck":"Food truck profit & taxes, tracked by event",
}

def tlen(d,t,f): return d.textlength(t,font=f)
def wrap(d,t,f,mw):
    words=t.split(); lines=[]; cur=""
    for w in words:
        s=(cur+" "+w).strip()
        if tlen(d,s,f)<=mw: cur=s
        else: lines.append(cur); cur=w
    if cur: lines.append(cur)
    return lines
def mark(d,x,y,s):
    d.rounded_rectangle([x,y,x+s,y+s],radius=int(s*0.22),fill=WHITE)
    pad=s*0.20; bw=s*0.15; gap=s*0.075; base=y+s-pad; bx=x+pad
    for h,c in zip([0.26,0.42,0.58],[TEAL_L,TEAL,NAVY]):
        d.rounded_rectangle([bx,base-s*h,bx+bw,base],radius=int(bw*0.25),fill=c); bx+=bw+gap
    r=s*0.18; cx,cy=x+s-pad*0.5,y+pad*0.6
    d.ellipse([cx-r,cy-r,cx+r,cy+r],fill=TEAL)
    d.line([cx-r*0.45,cy,cx-r*0.05,cy+r*0.4],fill=WHITE,width=4)
    d.line([cx-r*0.05,cy+r*0.4,cx+r*0.55,cy-r*0.45],fill=WHITE,width=4)
def check(d,x,y,s=44):
    d.ellipse([x,y,x+s,y+s],fill=TEAL)
    d.line([x+s*0.27,y+s*0.52,x+s*0.43,y+s*0.70],fill=WHITE,width=6)
    d.line([x+s*0.43,y+s*0.70,x+s*0.75,y+s*0.30],fill=WHITE,width=6)
def footer(d, line2=None):
    d.rectangle([0,H-170,W,H],fill=TEAL)
    mark(d,55,H-138,105)
    d.text((185,H-120),"LedgerLite",font=F(46),fill=WHITE)
    cta="→ Find it on Etsy"
    d.text((W-55-tlen(d,cta,F(38)),H-118),cta,font=F(38),fill=WHITE)

def product_pin(key,cfg):
    img=Image.new("RGB",(W,H),WHITE); d=ImageDraw.Draw(img)
    d.rectangle([0,0,W,720],fill=NAVY)
    mark(d,60,55,90); d.text((175,72),"LedgerLite",font=F(40),fill=WHITE)
    y=210
    for ln in wrap(d,HOOKS[key],F(66),880):
        d.text((60,y),ln,font=F(66),fill=WHITE); y+=82
    # dashboard card straddling boundary
    inc,exp,net,_=numbers(cfg)
    cx0,cy0,cw=60,600,W-120
    d.rounded_rectangle([cx0,cy0,cx0+cw,cy0+300],radius=28,fill=LIGHTER,outline=(203,213,224),width=3)
    d.text((cx0+40,cy0+30),"NET PROFIT",font=F(30),fill=TEAL)
    d.text((cx0+40,cy0+68),money(net),font=F(120),fill=TEAL)
    chip=(cw-120)//2
    for i,(lab,val,col) in enumerate([("Income",inc,GREEN),("Expenses",exp,RED)]):
        ccx=cx0+40+i*(chip+40)
        d.rounded_rectangle([ccx,cy0+198,ccx+chip,cy0+288],radius=14,fill=WHITE,outline=(203,213,224),width=2)
        d.text((ccx+22,cy0+208),lab,font=F(24,False),fill=GREY)
        d.text((ccx+22,cy0+242),money(val),font=F(34),fill=col)
    # benefits
    by=980
    bens=[f"Auto-fills your IRS {cfg.get('form','Schedule E')} summary",
          f"Profit by {cfg.get('entity_lower','property')} & by month",
          "One-time purchase — no subscription"]
    for b in bens:
        check(d,60,by); d.text((130,by+2),b,font=F(36,False),fill=INK); by+=86
    footer(d)
    img.save(os.path.join(OUT,f"pin-{key}.png")); print("  pin-"+key)

def bundle_pin(key,cfg):
    img=Image.new("RGB",(W,H),WHITE); d=ImageDraw.Draw(img)
    d.rectangle([0,0,W,560],fill=NAVY)
    mark(d,60,55,90); d.text((175,72),"LedgerLite",font=F(40),fill=WHITE)
    d.text((60,210),cfg["name"],font=F(70),fill=TEAL_L)
    d.text((60,320),f"{len(cfg['members'])} trackers — one purchase",font=F(46),fill=WHITE)
    pf=F(120); price_lab=f"${cfg['price']}"; d.text((60,400),price_lab,font=pf,fill=WHITE)
    al=f"${cfg['anchor']}"; d.text((60+tlen(d,price_lab,pf)+30,455),al,font=F(46,False),fill=(150,180,205))
    y=620
    show=cfg["members"][:8]
    for mk in show:
        nm=DISPLAY[mk][0].replace(" Bookkeeping","").replace(" Tracker","")
        check(d,60,y); d.text((130,y+2),nm,font=F(38,False),fill=INK); y+=80
    if len(cfg["members"])>8:
        d.text((130,y),f"...and {len(cfg['members'])-8} more",font=F(36),fill=TEAL); y+=70
    footer(d,"Schedule C & E · Excel & Google Sheets")
    img.save(os.path.join(OUT,f"pin-bundle-{key}.png")); print("  pin-bundle-"+key)

def shop_pin():
    img=Image.new("RGB",(W,H),NAVY); d=ImageDraw.Draw(img)
    mark(d,W//2-90,120,180)
    t="LedgerLite"; d.text((W//2-tlen(d,t,F(72))//2,330),t,font=F(72),fill=WHITE)
    for i,ln in enumerate(["Bookkeeping that does","your taxes for you"]):
        d.text((W//2-tlen(d,ln,F(56))//2,440+i*70),ln,font=F(56),fill=TEAL_L)
    trades=["Airbnb & rentals","Etsy & Amazon sellers","Freelancers & 1099","Real estate agents",
            "Photographers","Cleaners","Hair & beauty","Rideshare drivers","Food trucks"]
    y=640
    for t in trades:
        check(d,180,y); d.text((250,y+2),t,font=F(36,False),fill=WHITE); y+=82
    d.rounded_rectangle([180,H-250,W-180,H-150],radius=20,fill=TEAL)
    cta="Pick your trade →"; d.text((W//2-tlen(d,cta,F(44))//2,H-225),cta,font=F(44),fill=WHITE)
    img.save(os.path.join(OUT,"pin-shop.png")); print("  pin-shop")

if __name__=="__main__":
    print("pins:")
    for k,c in VARIANTS.items(): product_pin(k,c)
    for k,c in BUNDLES.items(): bundle_pin(k,c)
    shop_pin()
    print("done ->", os.path.relpath(OUT))
