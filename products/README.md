# Product Catalog — Schedule E Bookkeeping Trackers

Niche, formula-driven Google Sheets / Excel templates sold as **instant digital downloads**
on Etsy + Gumroad/Payhip. Chosen on market data (see `../DECISION.md`): no domain edge
required, a few hours to produce, defensible (formulas = genuine human authorship, not AI
images), high price point with price-insensitive buyers, fully passive delivery.

## Products

| Folder | Product | Tax form | Target buyer | Keyword | Price |
|--------|---------|----------|--------------|---------|-------|
| `str-tracker/` | Short-Term Rental Tracker | Schedule E | Airbnb/VRBO hosts | "airbnb spreadsheet" | $19 → $29 |
| `landlord-tracker/` | Rental Property Tracker | Schedule E | Long-term landlords | "rental property spreadsheet" | $19 → $29 |
| `seller-tracker/` | Online Seller Bookkeeping Tracker | Schedule C | Etsy/Amazon sellers | "etsy bookkeeping spreadsheet" | $19 → $29 |
| `freelancer-tracker/` | Freelancer & Self-Employed Tracker | Schedule C | Freelancers / 1099 | "freelancer bookkeeping spreadsheet" | $19 → $29 |
| `realtor-tracker/` | Real Estate Agent Tracker | Schedule C | Realtors / brokers | "real estate agent spreadsheet" | $24 → $39 |
| `photographer-tracker/` | Photographer Tracker | Schedule C | Photographers | "photographer bookkeeping spreadsheet" | $19 → $29 |

**Bundles** (see `STOREFRONT.md`): Rental Investor $39 · Self-Employed $39 · Everything $69.
**Launch:** step-by-step printable in `LAUNCH-CHECKLIST.md`.
`STOREFRONT.md` has store-name ideas, the store bio/About, and cross-sell snippets.
`QUICKSTART-EMAIL.md` has the reusable post-purchase quick-start + review email (drives reviews,
which drive early ranking).

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
- **Seller:** income 2710 / expenses 588.60 / net 2121.40 (Etsy Shop net 655.40)
- **Freelancer:** income 4350 / expenses 637.50 / net 3712.50 (Acme Co net 1800, overhead −237.50)
- **Realtor:** income 16500 / expenses 5530 / net 10970 (123 Main St net 6275, overhead −555)
- **Photographer:** income 3470 / expenses 1043 / net 2427 (Smith Wedding net 2400, overhead −565)

Open in Excel or Google Sheets and the live formulas reproduce these. *(Note: this environment's
LibreOffice headless can't recalc, so the math is verified by logic-equivalence — do open the
file once in Sheets before listing.)*

## Add another variant (fastest way to grow)
In `build_tracker.py`, copy a `VARIANTS` entry and change the fields. The engine is fully
parameterized — even the **tax form and entity labels** are configurable, so you are not
limited to Schedule E. Per-variant keys:
- Core: `dir`, `output`, `title`, `subtitle`, `income`, `expenses`, `placeholders`,
  `setup_col3`, `payee_header`, tip / `log_right` text, `samples`
- Tax-form/labeling (optional, default = Schedule E / "property"): `form`, `entity`,
  `entity_lower`, `entity_plural_lower`, `summary_sheet`, `summary_part`,
  `income_line_label`, `expense_total_label`, `net_label`, `setup_col2`, `setup_col4`,
  `setup_col4_currency`

The `seller` variant (Schedule C, entity = "Shop") is the worked example of using those
optional keys. Good next candidates (same engine):
- **Single-trade bookkeeping** (cleaner, photographer, salon, food truck) — Schedule C, swap categories
- **Single-property simple** version (set `MAX_PROPS = 1`)
- **Freelancer / 1099 contractor** tracker — Schedule C

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
