#!/usr/bin/env python3
"""
Build the customer-facing Setup Guide PDF for each tracker variant.

Reads the same VARIANTS config as build_tracker.py, so the guide always matches the
workbook it ships with.

Usage:
  python3 build_guide.py            # build all guides
  python3 build_guide.py str        # build one
"""
import os, sys
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, HRFlowable, ListFlowable, ListItem)

from build_tracker import VARIANTS

NAVY = colors.HexColor("#1F3A5F"); TEAL = colors.HexColor("#2C7A7B")
GREY = colors.HexColor("#4A5568"); LIGHT = colors.HexColor("#EBF2F7")

styles = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=styles["Heading1"], textColor=NAVY, fontSize=22, spaceAfter=4, leading=26)
SUB = ParagraphStyle("SUB", parent=styles["Normal"], textColor=GREY, fontSize=10.5, spaceAfter=14, leading=14)
H2 = ParagraphStyle("H2", parent=styles["Heading2"], textColor=TEAL, fontSize=14, spaceBefore=14, spaceAfter=6, leading=17)
BODY = ParagraphStyle("BODY", parent=styles["Normal"], fontSize=10.5, leading=15, textColor=colors.HexColor("#1A202C"), spaceAfter=6)
BULLET = ParagraphStyle("BULLET", parent=BODY, leftIndent=6, spaceAfter=3)
SMALL = ParagraphStyle("SMALL", parent=styles["Normal"], fontSize=8.5, leading=11, textColor=GREY)


def bullets(items):
    return ListFlowable(
        [ListItem(Paragraph(t, BULLET), leftIndent=12, value="•") for t in items],
        bulletType="bullet", start="•", leftIndent=10)


def build_guide(cfg):
    name = cfg["title"].strip().split("  ")[0]            # e.g. "Short-Term Rental"
    noun = cfg["noun"]
    story = []
    story.append(Paragraph(name + " Tracker", H1))
    story.append(Paragraph("Setup &amp; User Guide &nbsp;·&nbsp; " + cfg["subtitle"].strip(), SUB))
    story.append(HRFlowable(width="100%", color=LIGHT, thickness=2, spaceAfter=12))

    story.append(Paragraph("What you bought", H2))
    story.append(Paragraph(
        f"A single, formula-driven spreadsheet that does your {noun} bookkeeping for you. "
        "Log each transaction once, and the workbook automatically organizes everything into a "
        "tax-ready <b>IRS Schedule E</b> summary and a live profit dashboard — per property. "
        "No subscriptions, no logins, no monthly fees. It's yours forever.", BODY))

    story.append(Paragraph("The 6 tabs", H2))
    tab_rows = [
        ["Tab", "What it's for"],
        ["Start Here", "A one-page overview and tips. Read this first."],
        ["Setup", "Type your property names (up to 5). Everything else links to them automatically."],
        ["Transactions", "Your logbook. Add one row per income or expense. The only tab you touch daily."],
        ["Categories", "Reference list of every category and its matching Schedule E line. Feeds the dropdowns."],
        ["Schedule E Summary", "Fills itself in. Every expense mapped to the correct Schedule E line, per property."],
        ["Dashboard", "Live totals: net profit, income vs. expenses, by property and by month."],
    ]
    t = Table(tab_rows, colWidths=[1.5*inch, 4.6*inch])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), TEAL), ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"), ("FONTNAME", (0,1), (0,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9.5), ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, LIGHT]),
        ("GRID", (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ("TOPPADDING", (0,0), (-1,-1), 6), ("BOTTOMPADDING", (0,0), (-1,-1), 6),
        ("LEFTPADDING", (0,0), (-1,-1), 8),
    ]))
    story.append(t)

    story.append(Paragraph("Get started in 4 steps", H2))
    story.append(bullets([
        "<b>1. Name your properties.</b> Open the <b>Setup</b> tab and type each property name in column B "
        "(replace the sample names). You can have up to 5.",
        "<b>2. Delete the sample data.</b> The <b>Transactions</b> tab has 8 example rows so you can see how "
        "it works. Once you understand it, delete those rows (select rows 4–11, right-click, Delete) before "
        "adding your own.",
        "<b>3. Log a transaction.</b> On <b>Transactions</b>, fill one row: pick the Date, choose the "
        "Property and Type (Income / Expense) from the dropdowns, pick a Category, enter the Amount. The "
        "Month column fills in by itself.",
        "<b>4. Read your reports.</b> Open <b>Schedule E Summary</b> and <b>Dashboard</b> any time — they "
        "update the instant you add a row. Nothing to refresh.",
    ]))

    story.append(Paragraph("How to log it the right way (important)", H2))
    story.append(bullets(cfg["log_right"]))

    story.append(Paragraph("Use it in Excel or Google Sheets", H2))
    story.append(Paragraph(
        "<b>Excel / Numbers:</b> just open the .xlsx file.<br/>"
        "<b>Google Sheets:</b> go to Google Drive → New → File upload → select the .xlsx. Then open it and "
        "choose <i>File → Save as Google Sheets</i>. All formulas, dropdowns and reports carry over.", BODY))

    story.append(Paragraph("Each tax year", H2))
    story.append(Paragraph(
        f"Keep one file per year. At the start of a new year, make a copy and rename it (e.g. "
        f"“{name} Tracker 2027”), then clear the Transactions rows. Your prior year stays "
        "untouched for your records.", BODY))

    story.append(Spacer(1, 10))
    story.append(HRFlowable(width="100%", color=LIGHT, thickness=2, spaceAfter=8))
    story.append(Paragraph(
        "<b>Disclaimer.</b> This template is a bookkeeping organizer, not tax, legal, or accounting advice. "
        "The Schedule E line mapping is provided for convenience only. Tax rules vary by situation and "
        "jurisdiction — always confirm your specific filing with a qualified tax professional before "
        "submitting. Made with spreadsheet software; design and formulas by the seller.", SMALL))

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), cfg["dir"], "Setup-Guide.pdf")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    SimpleDocTemplate(out, pagesize=letter, topMargin=0.7*inch, bottomMargin=0.7*inch,
                      leftMargin=0.8*inch, rightMargin=0.8*inch,
                      title=name + " Tracker — Setup Guide").build(story)
    print("wrote", out)


if __name__ == "__main__":
    keys = sys.argv[1:] or list(VARIANTS)
    for k in keys:
        if k not in VARIANTS:
            print("unknown variant:", k); continue
        build_guide(VARIANTS[k])
