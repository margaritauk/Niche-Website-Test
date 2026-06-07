# Social Launch Kit — LedgerLite (Pinterest + Instagram)

**Why Pinterest first:** it's a search engine, not a feed — pins keep driving traffic for months,
and it's the #1 source for spreadsheet/printable shops. Instagram/TikTok are secondary.

**Assets:** vertical pins (1000×1500) are in `pins/` — one per product, per bundle, and a shop pin.
Each pin **links to its listing** (paste your Etsy/Gumroad URL when you upload the pin).

---

## Pinterest setup (one-time, ~15 min)
1. Create a **free Pinterest Business account** (pinterest.com/business) → name it **LedgerLite**.
2. Profile photo: `shop/assets/icon-500.png`. Bio: *"Bookkeeping spreadsheets that do your taxes for you. For hosts, sellers & the self-employed. Excel & Google Sheets."*
3. **Claim your website** (your Gumroad/Etsy or domain) for analytics + attribution.
4. Make **boards** (these double as keywords):
   `Small Business Bookkeeping` · `Etsy Seller Tips` · `Airbnb & Rental Tips` · `Freelance & Self-Employed` · `Tax Prep & Organization` · `Side Hustle Tools` · `Real Estate Agent Tips` · `Gig Economy / Driver Tips`

## Posting cadence (keep it simple)
- **5–10 pins/day** is ideal but unrealistic by hand — do **1–3/day**, consistently.
- Schedule them free inside Pinterest (Create → schedule date) or with Tailwind.
- Re-pin the same image to **multiple relevant boards** (space them a few days apart).
- Fresh pins win: over time, make 2–3 image variants per product (re-run `build_pins.py` after editing a hook).

---

## Per-pin copy (Title + Description + Board)
Pinterest titles ≤ 100 chars; descriptions ≤ 500 chars, keyword-rich, end with a soft CTA.
Use this hashtag set on most pins:
`#bookkeeping #smallbusiness #selfemployed #taxprep #spreadsheet #sidehustle #etsyseller #googlesheets`

**Short-Term Rental** → board: Airbnb & Rental Tips
- *Title:* Airbnb Bookkeeping Spreadsheet | Short-Term Rental Tax Tracker
- *Desc:* Hosting on Airbnb or VRBO? Track every booking and expense in one simple spreadsheet and get an instant, tax-ready IRS Schedule E summary. Works in Excel & Google Sheets. One-time purchase, no subscription. Tap to get it. #airbnbhost #rentalproperty #bookkeeping #taxprep

**Rental Property** → Airbnb & Rental Tips
- *Title:* Landlord Spreadsheet | Rental Property Income, Expense & Tax Tracker
- *Desc:* The simple way for landlords to track rent, expenses and mortgage interest — with an automatic IRS Schedule E summary. Handles deposits & multiple properties. Excel & Google Sheets, one-time purchase. #landlord #rentalproperty #realestateinvesting #bookkeeping

**Online Seller** → Etsy Seller Tips
- *Title:* Etsy Seller Bookkeeping Spreadsheet | See Your Real Profit After Fees
- *Desc:* Selling on Etsy or Amazon? Stop guessing your profit. Log sales and expenses once and see your REAL profit after marketplace fees — plus an auto IRS Schedule C summary. Separates fees and COGS for you. #etsyseller #smallbusiness #bookkeeping #etsytips

**Freelancer / 1099** → Freelance & Self-Employed
- *Title:* Freelancer Bookkeeping Spreadsheet | Profit by Client + Schedule C
- *Desc:* Freelancer or 1099 contractor? Track income and expenses by client, know your real profit, and set aside the right amount for quarterly taxes. Auto IRS Schedule C summary. Excel & Google Sheets. #freelancer #selfemployed #1099 #taxprep

**Real Estate Agent** → Real Estate Agent Tips
- *Title:* Realtor Bookkeeping Spreadsheet | Commission, Mileage & Tax Tracker
- *Desc:* Agents: see your true take-home after your brokerage split, track mileage and dues, and get an automatic IRS Schedule C summary. Built for the deductions agents miss. #realestateagent #realtor #realtorlife #bookkeeping

**Photographer** → Small Business Bookkeeping
- *Title:* Photographer Bookkeeping Spreadsheet | Profit by Shoot + Schedule C
- *Desc:* Photographers: see which shoots actually pay. Track session fees, print sales, gear and editing costs — with an auto IRS Schedule C summary. Excel & Google Sheets, one-time purchase. #photographer #photographybusiness #smallbusiness #bookkeeping

**Cleaning Business** → Small Business Bookkeeping
- *Title:* Cleaning Business Spreadsheet | Income, Expense & Tax Tracker
- *Desc:* House cleaners: track jobs, tips, supplies and mileage with an automatic IRS Schedule C summary and profit by client. No accounting know-how needed. #cleaningbusiness #cleaningbusinesstips #smallbusiness #bookkeeping

**Hair Stylist & Beauty** → Small Business Bookkeeping
- *Title:* Hair Stylist Bookkeeping Spreadsheet | Booth Renter Tax Tracker
- *Desc:* Stylists, barbers & nail techs: track services, retail, tips and booth rent — with an auto IRS Schedule C summary. Built for booth & suite renters. Excel & Google Sheets. #hairstylist #boothrenter #estheticism #salonowner

**Rideshare & Delivery** → Gig Economy / Driver Tips
- *Title:* Uber & DoorDash Spreadsheet | Mileage, Income & Tax Tracker
- *Desc:* Drive for Uber, Lyft or DoorDash? See which app pays best, track mileage (your biggest deduction) and get an auto IRS Schedule C summary. Plain-English mileage guide included. #doordash #uberdriver #gigeconomy #mileagetracker

**Food Truck & Vendor** → Small Business Bookkeeping
- *Title:* Food Truck Spreadsheet | Profit by Event + COGS & Tax Tracker
- *Desc:* Food trucks & market vendors: see which events make money, track food costs (COGS), commissary rent and permits — with an auto IRS Schedule C summary. #foodtruck #foodtruckbusiness #smallbusiness #bookkeeping

**Everything Bundle** → Side Hustle Tools
- *Title:* 10 Small Business Bookkeeping Spreadsheets | Tax Tracker Bundle
- *Desc:* Every LedgerLite tracker in one bundle — rentals, Etsy, freelance, real estate, photography, cleaning, beauty, rideshare & food trucks. Schedule C & E ready. $69 for all ten. #bookkeeping #smallbusiness #sidehustle #taxprep

**Shop pin** → pin to all boards
- *Title:* Bookkeeping Spreadsheets That Do Your Taxes For You
- *Desc:* LedgerLite makes simple, tax-ready bookkeeping spreadsheets for hosts, sellers, freelancers and trades. Log income & expenses once — get an instant IRS summary. Excel & Google Sheets. Pick your trade → #bookkeeping #smallbusiness #selfemployed #sidehustle

---

## Instagram / TikTok (secondary)
- **Bio:** `LedgerLite 📊 Bookkeeping that does your taxes for you · Excel & Google Sheets · Shop ↓`
- **Profile pic:** `shop/assets/icon-500.png`. Use the **pins** as carousel/feed posts too (or the square product images from each `images/` folder).
- **Caption template:**
  > Doing your own books shouldn't be this hard. 😮‍💨 This [trade] bookkeeping spreadsheet sorts every dollar onto the right tax line and shows your real profit — by [client/property/shop] and by month. One-time purchase, works in Excel & Google Sheets. Link in bio. 💛
  > .
  > #[trade] #smallbusiness #bookkeeping #selfemployed #taxprep #sidehustle #spreadsheet #googlesheets
- **Reel idea (15s):** screen-record scrolling the tracker → typing one expense → the Dashboard number updating. Caption: "POV: your bookkeeping does itself."

## Launch week plan
- **Day 1:** publish the shop pin + Everything Bundle pin; IG intro post.
- **Days 2–10:** one product pin/day (start with your best-keyword products), each linked to its listing, pinned to its board + re-pinned to a second board.
- **Ongoing:** 1–3 pins/day, mixing products + bundles; repost winners.
