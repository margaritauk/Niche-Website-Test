#!/usr/bin/env python3
"""
Assemble the bundle listings: hero image + 'what's included' collage + a ready-to-upload
ZIP (all member trackers + guides + an index) for each bundle.

Output -> bundles/<bundle>/  :  hero.png, included.png, <bundle>.zip

Usage: python3 build_bundles.py
"""
import os, zipfile, textwrap
from PIL import Image, ImageDraw, ImageFont
from build_tracker import VARIANTS
from build_mockups import DISPLAY

NAVY=(31,58,95); TEAL=(44,122,123); TEAL_L=(72,160,162); WHITE=(255,255,255)
LIGHT=(235,242,247); GREY=(113,128,150); INK=(26,32,44)
FROOT="/usr/share/fonts/truetype/dejavu"
def F(s,b=True): return ImageFont.truetype(f"{FROOT}/DejaVuSans{'-Bold' if b else ''}.ttf", s)

HERE = os.path.dirname(os.path.abspath(__file__))
BRAND = "LedgerLite"

BUNDLES = {
    "rental-investor": {
        "name": "Rental Investor Bundle", "price": 39, "anchor": 58,
        "members": ["str", "landlord"],
        "blurb": "For Airbnb hosts who also rent long-term — both Schedule E trackers.",
    },
    "self-employed": {
        "name": "Self-Employed Bundle", "price": 39, "anchor": 58,
        "members": ["seller", "freelancer"],
        "blurb": "Sell online and freelance? Both Schedule C trackers, one price.",
    },
    "everything": {
        "name": "Everything Bundle", "price": 69, "anchor": 220,
        "members": list(VARIANTS.keys()),
        "blurb": "All ten trackers — every business, both tax forms.",
    },
}

def tlen(d,t,f): return d.textlength(t,font=f)
def pill(d,x,y,label,font,fg=WHITE,bg=TEAL,padx=28,pady=15):
    w=tlen(d,label,font); h=font.getbbox("Ay")[3]
    d.rounded_rectangle([x,y,x+w+2*padx,y+h+2*pady],radius=(h+2*pady)//2,fill=bg)
    d.text((x+padx,y+pady-2),label,font=font,fill=fg); return x+w+2*padx
def mark(d,x,y,s):
    d.rounded_rectangle([x,y,x+s,y+s],radius=int(s*0.22),fill=WHITE)
    pad=s*0.20; bw=s*0.15; gap=s*0.075; base=y+s-pad; bx=x+pad
    for h,c in zip([0.26,0.42,0.58],[TEAL_L,TEAL,NAVY]):
        d.rounded_rectangle([bx,base-s*h,bx+bw,base],radius=int(bw*0.25),fill=c); bx+=bw+gap
    r=s*0.18; cx,cy=x+s-pad*0.5,y+pad*0.6
    d.ellipse([cx-r,cy-r,cx+r,cy+r],fill=TEAL)
    d.line([cx-r*0.45,cy,cx-r*0.05,cy+r*0.4],fill=WHITE,width=max(3,int(s*0.03)))
    d.line([cx-r*0.05,cy+r*0.4,cx+r*0.55,cy-r*0.45],fill=WHITE,width=max(3,int(s*0.03)))

def hero(key, cfg, out):
    S=2000; img=Image.new("RGB",(S,S),NAVY); d=ImageDraw.Draw(img)
    mark(d,120,130,220)
    d.text((390,150),BRAND,font=F(64),fill=WHITE)
    # bundle name (wrap)
    name=cfg["name"]; nf=F(104)
    if tlen(d,name,nf)>S-470: nf=F(86)
    d.text((390,250),name,font=nf,fill=TEAL_L)
    mem_keys=cfg["members"]; small=len(mem_keys)<=4
    y=560
    d.text((120,470),f"{len(mem_keys)} trackers — one purchase",font=F(58),fill=WHITE)
    lf=F(50,False) if small else F(46,False)
    for mk in mem_keys[:10]:
        nm=DISPLAY[mk][0].replace(" Bookkeeping","").replace(" Tracker","")
        d.ellipse([130,y+8,178,y+56],fill=TEAL); d.line([143,y+33,158,y+48],fill=WHITE,width=7)
        d.line([158,y+48,176,y+18],fill=WHITE,width=7)
        d.text((210,y),nm,font=lf,fill=WHITE)
        if small:
            d.text((210,y+62),DISPLAY[mk][1],font=F(38,False),fill=(180,205,224)); y+=170
        else:
            y+=100 if len(mem_keys)>6 else 120
    if small:
        # benefit strip to fill the space
        by0=1180
        for i,bn in enumerate(["Both tax-ready, side by side","One-time purchase — no subscription",
                               "Works in Excel & Google Sheets"]):
            d.ellipse([130,by0+i*120+6,178,by0+i*120+54],fill=TEAL)
            d.line([143,by0+i*120+31,158,by0+i*120+46],fill=WHITE,width=7)
            d.line([158,by0+i*120+46,176,by0+i*120+16],fill=WHITE,width=7)
            d.text((210,by0+i*120),bn,font=F(44,False),fill=WHITE)
    # price band
    by=1760
    d.rounded_rectangle([120,by,S-120,by+150],radius=24,fill=TEAL)
    d.text((160,by+38),"Save vs. buying separately",font=F(46,False),fill=WHITE)
    pf=F(72); lab=f"${cfg['price']}"
    anchor_lab=f"${cfg['anchor']}"
    d.text((S-180-tlen(d,lab,pf),by+30),lab,font=pf,fill=WHITE)
    d.text((S-180-tlen(d,lab,pf)-tlen(d,anchor_lab,F(48))-30,by+52),
           anchor_lab,font=F(48,False),fill=(180,205,224))
    img.save(out); print("  hero ->", os.path.relpath(out))

def collage(key, cfg, out):
    S=2000; img=Image.new("RGB",(S,S),WHITE); d=ImageDraw.Draw(img)
    d.rectangle([0,0,S,300],fill=NAVY)
    mark(d,70,70,160); d.text((280,95),f"{len(cfg['members'])} trackers inside",font=F(76),fill=WHITE)
    d.text((280,205),f"{BRAND}  ·  {cfg['name']}",font=F(40,False),fill=(180,205,224))
    members=cfg["members"]
    show=members[:6] if len(members)>6 else members
    cols=2 if len(show)<=4 else 3
    rows=(len(show)+cols-1)//cols
    pad=60; gw=(S-pad*(cols+1))//cols; gh=(S-340-pad*(rows+1))//rows
    for i,m in enumerate(show):
        r,c=divmod(i,cols); x=pad+c*(gw+pad); y=340+pad+r*(gh+pad)
        thumb=Image.open(os.path.join(HERE,VARIANTS[m]["dir"],"images","01-hero.png")).resize((gw,gw),Image.LANCZOS)
        # crop top portion (hero header) to gh
        thumb=thumb.crop((0,0,gw,min(gh,gw)))
        img.paste(thumb,(x,y))
        d.rounded_rectangle([x,y,x+gw,y+min(gh,gw)],radius=2,outline=(203,213,224),width=3)
    if len(members)>6:
        d.text((pad,S-90),f"...and {len(members)-6} more",font=F(46),fill=TEAL)
    img.save(out); print("  collage ->", os.path.relpath(out))

def make_zip(key, cfg, out):
    readme = [f"{BRAND} — {cfg['name']}", "="*40, "",
              "Thanks for your purchase! This bundle includes the following trackers,",
              "each with its own Setup Guide (PDF). Open a guide first — setup takes ~5 min.",
              "", "Included:"]
    with zipfile.ZipFile(out,"w",zipfile.ZIP_DEFLATED) as z:
        for m in cfg["members"]:
            cfgm=VARIANTS[m]; folder=DISPLAY[m][0]
            xlsx=os.path.join(HERE,cfgm["dir"],cfgm["output"])
            pdf=os.path.join(HERE,cfgm["dir"],"Setup-Guide.pdf")
            z.write(xlsx, f"{folder}/{cfgm['output']}")
            z.write(pdf,  f"{folder}/Setup-Guide.pdf")
            readme.append(f"  - {folder}")
        readme += ["", "Tip: works in Excel, Google Sheets & Numbers. For Sheets, upload the",
                   ".xlsx to Google Drive, open it, and choose File > Save as Google Sheets.",
                   "", "Not tax advice — confirm your filing with a qualified professional."]
        z.writestr("READ-ME-FIRST.txt", "\n".join(readme))
    print("  zip  ->", os.path.relpath(out), f"({os.path.getsize(out)//1024} KB)")

if __name__ == "__main__":
    for key,cfg in BUNDLES.items():
        d=os.path.join(HERE,"bundles",key); os.makedirs(d,exist_ok=True)
        print(cfg["name"]+":")
        hero(key,cfg,os.path.join(d,"hero.png"))
        collage(key,cfg,os.path.join(d,"included.png"))
        make_zip(key,cfg,os.path.join(d,f"{key}-bundle.zip"))
