# Dokan Khata

A stock, sales and cash app for shops of any size. Works on phones, tablets, laptops and PCs, installs like an app (PWA), and keeps working offline. English and Bangla.

## What it does
- **Products**: name (English and Bangla), photo, price, cost, stock, unit, category, barcode, low-stock alert. Add stock with one tap and optionally log the purchase.
- **Sell**: tap pictures to build a bill, or scan a barcode (USB scanner, typed code, or phone camera on Chrome/Android). Cash, bKash, Nagad, card or due (baki). Stock drops automatically.
- **Money**: sales by day (tap to view or delete, which returns stock), income and expense book, list of who owes you money with "mark as paid".
- **Home**: sales, profit after product cost, running costs, net earnings, daily sales chart, best sellers, low stock, stock value, total due.
- **My shop**: what customers see. They pick items and send the order to your WhatsApp.
- **Setup**: shop name, WhatsApp number, language, backup download/restore, sales CSV export.

## Plans and access codes
- Plans: Shuru ৳350, Bepsha ৳500, Pro ৳1000 per month. Bepsha has everything except the online shop. Edit `TIERS` in `index.html` to change what each plan includes.
- First open shows the Terms and conditions (English and Bangla), then a screen explaining the free 10-minute trial. After 10 minutes the app locks until an access code is entered. Change `demoSeconds` in `CONFIG` to adjust.
- Access codes are signed with a private key. Create them on `admin.html` after loading `admin-key.json` (kept privately, never in this repo). A code carries the business name, plan and end date, and arrives as a link the business taps.
- Payments: bKash, Nagad or Rocket send-money. The customer sends the Transaction ID through WhatsApp; you check it and send a code. Set your numbers in `CONFIG` in `index.html` (`ownerWhatsApp`, `pay`).
- Limit: the app runs fully in the browser, so a technical person could bypass the lock by editing the page. Server-side checks come with the login and sync work.

## Google Apps Script version
`apps-script/` runs the same app from a Google Sheet: data for every shop is saved in the sheet, access codes and plan payments are managed there, and the owner gets an email for each payment. See `apps-script/SETUP.md`. After changing `index.html`, run `python3 apps-script/build.py` to rebuild `apps-script/Index.html`.

## How it is built
One `index.html` (plain HTML, CSS, JavaScript, no build step), `manifest.webmanifest`, `sw.js` for offline, and `icons/`. Data is stored in the browser's IndexedDB on each device.

## Run it locally
```
python3 -m http.server 8000
```
Then open http://localhost:8000.

## Put it online for free
Any static host works. Options with free plans: GitHub Pages, Cloudflare Pages, Netlify. Upload this folder as-is; HTTPS is required for install and offline.

## Not built yet (needs online accounts)
- Sign-in and sync between devices.
- A public shop link that customers can open (today, "My shop" is a preview on the owner's device).
- Staff accounts, multiple branches, subscription billing (bKash/SSLCommerz).
Planned backend: Supabase free tier (database + sign-in), which keeps running costs near zero until there are many shops.
