# Short-Term Rental (Airbnb/VRBO) Bookkeeping & Tax Tracker

The first product: a niche, formula-driven Google Sheets / Excel template sold as an
**instant digital download** on Etsy + Gumroad/Payhip. Chosen on market data (see
`../DECISION.md`) for the best fit with: no domain edge required, only a few hours to
produce, defensible (formulas = genuine human authorship, not AI images), high price
point with price-insensitive buyers, and fully passive delivery.

## What's in this folder

| File | What it is | For |
|------|-----------|-----|
| `Short-Term-Rental-Tracker.xlsx` | The product itself — 6-tab automated workbook | **Upload to store** |
| `Setup-Guide.pdf` | Customer-facing setup & user guide | **Upload to store** |
| `listing-copy.md` | Titles, 13 Etsy tags, full description, pricing, mockup shot list | Copy/paste when listing |
| `build_tracker.py` | Generator for the .xlsx (re-run to edit or spin up variants) | Source |
| `build_guide.py` | Generator for the PDF guide | Source |

## The product (what the buyer gets)

A single spreadsheet that:
- Logs income & expenses per property (up to 5) via dropdown-driven categories
- Auto-maps every expense to its **IRS Schedule E** line
- Produces a tax-ready Schedule E summary per property
- Shows a live dashboard: net profit, income vs. expenses, by-property and by-month

Six tabs: **Start Here · Setup · Transactions · Categories · Schedule E Summary · Dashboard**.
Pre-filled with 8 example rows so buyers see how it works, then delete and start.

## Regenerate / edit

```bash
pip install openpyxl reportlab
python3 build_tracker.py   # -> Short-Term-Rental-Tracker.xlsx
python3 build_guide.py     # -> Setup-Guide.pdf
```

The numbers were verified by reimplementing the workbook's SUMIFS against the sample
rows (Total income 2205 / expenses 390 / net 1815; Beach Cottage net 1061.50). Open in
Excel or Google Sheets and the live formulas reproduce these.

## Spinning up sibling variants (same engine, swap labels)

The fastest way to grow the catalog without new builds. In `build_tracker.py`, edit
`EXPENSE_CATEGORIES` / `INCOME_CATEGORIES` and the banner text to produce:
- **Long-term landlord** version (rename platform fields, drop occupancy tax)
- **Single-trade bookkeeping** (cleaner, photographer, salon) — swap categories
- **Single-property simple** version (set `MAX_PROPS = 1`)

Each becomes a new listing targeting its own keyword, then bundle them.

## Launch checklist

1. [ ] Validate the keyword in eRank/Everbee free tier ("airbnb spreadsheet", "str tracker")
2. [ ] Open the .xlsx in Google Sheets once to confirm formulas/dropdowns carry over
3. [ ] Make 5 product mockups (shot list in `listing-copy.md`)
4. [ ] List on Gumroad/Payhip (you own the buyer email) + Etsy (discovery)
5. [ ] Paste title/tags/description from `listing-copy.md`; keep the AI-disclosure line
6. [ ] Price at $19 intro, raise to $24–29 after 5–10 reviews
7. [ ] Add 2–3 sibling variants over the following weeks; bundle

## Important notes

- **Legal defensibility:** functional formulas/structure are human-authored work, which
  sidesteps the 2026 *Thaler v. Perlmutter* gap that leaves pure-AI images uncopyrightable.
- **Etsy AI disclosure (2026):** mandatory — the disclosure sentence is already in the
  description; list under "Designed by," not "Made by."
- **Not tax advice:** the template is a bookkeeping organizer; the disclaimer is in the
  workbook, the PDF, and the listing.
