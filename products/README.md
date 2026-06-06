# Product Catalog — Schedule E Bookkeeping Trackers

Niche, formula-driven Google Sheets / Excel templates sold as **instant digital downloads**
on Etsy + Gumroad/Payhip. Chosen on market data (see `../DECISION.md`): no domain edge
required, a few hours to produce, defensible (formulas = genuine human authorship, not AI
images), high price point with price-insensitive buyers, fully passive delivery.

## Products

| Folder | Product | Target buyer | Keyword | Price |
|--------|---------|--------------|---------|-------|
| `str-tracker/` | Short-Term Rental Tracker | Airbnb/VRBO hosts | "airbnb spreadsheet" | $19 → $29 |
| `landlord-tracker/` | Rental Property Tracker | Long-term landlords | "rental property spreadsheet" | $19 → $29 |

**Bundle both for $39** ("Rental Investor Bundle") once they're live.

Each product folder contains:
- `*.xlsx` — the product (6-tab automated workbook) → **upload to store**
- `Setup-Guide.pdf` — customer-facing guide → **upload to store**
- `listing-copy.md` — titles, 13 Etsy tags, full description, pricing
- `mockups.md` — ready-to-use copy for the 5 product images

## The shared engine

Both products are generated from one parameterized source — add a product by adding a dict
to `VARIANTS` in `build_tracker.py`; nothing else changes.

```bash
pip install openpyxl reportlab

python3 build_tracker.py            # build all workbooks
python3 build_tracker.py str        # build one (str | landlord)
python3 build_guide.py              # build all setup guides (reads the same VARIANTS)
```

`build_guide.py` imports `VARIANTS` from `build_tracker.py`, so a guide can never drift from
its workbook. Output lands in each product's folder automatically (`cfg["dir"]`).

### What each workbook does
Six tabs — **Start Here · Setup · Transactions · Categories · Schedule E Summary · Dashboard**:
- Log income & expenses per property (up to 5) via dropdown-driven categories
- Auto-map every expense to its **IRS Schedule E** line (SUMIFS)
- Tax-ready Schedule E summary per property + live profit/cash-flow dashboard
- Ships with 8 sample rows so buyers see it working, then delete and start

### Verification
Numbers are verified by reimplementing each workbook's SUMIFS against its sample rows:
- **STR:** income 2205 / expenses 390 / net 1815 (Beach Cottage net 1061.50)
- **Landlord:** income 3575 / expenses 1395 / net 2180 (123 Oak St net 865)

Open in Excel or Google Sheets and the live formulas reproduce these. *(Note: this environment's
LibreOffice headless can't recalc, so the math is verified by logic-equivalence — do open the
file once in Sheets before listing.)*

## Add another variant (fastest way to grow)
In `build_tracker.py`, copy a `VARIANTS` entry and change: `dir`, `output`, `title`,
`subtitle`, `income`, `expenses`, `placeholders`, `setup_col3`, `payee_header`, the tip /
`log_right` text, and `samples`. Good next candidates (same engine):
- **Single-trade bookkeeping** (cleaner, photographer, salon) — swap categories, Schedule C
- **Single-property simple** version (set `MAX_PROPS = 1`)
- **Etsy / online-seller bookkeeping** — sales, fees, COGS, sales tax

## Launch checklist (per product)
1. [ ] Validate the keyword in eRank/Everbee free tier
2. [ ] Open the `.xlsx` in Google Sheets once — confirm formulas/dropdowns carry over
3. [ ] Make 5 mockups from `mockups.md`
4. [ ] List on Gumroad/Payhip (you own the buyer email) + Etsy (discovery)
5. [ ] Paste title/tags/description from `listing-copy.md`; keep the AI-disclosure line
6. [ ] Price at $19 intro → $24–29 after 5–10 reviews; add the $39 bundle

## Guardrails
- **Defensibility:** functional formulas/structure are human-authored, sidestepping the 2026
  *Thaler v. Perlmutter* gap that leaves pure-AI images uncopyrightable.
- **Etsy AI disclosure (2026):** mandatory — the disclosure line is in every listing; list
  under "Designed by," not "Made by."
- **Not tax advice:** the disclaimer is in each workbook, PDF, and listing.
