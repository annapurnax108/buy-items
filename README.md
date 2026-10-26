# Buy Items (PWA + Apps Script)

Same setup as the Sheet To-Do app. One shopping list per store, tracks how often each item is bought, and shows an estimated bill.

## 1. Sheet + backend
1. Open the Google Sheet. Extensions > Apps Script. Paste `Code.gs`, save.
2. Select `setup` in the function dropdown, Run, authorize. The execution log shows your **secret token** - copy it.
   This creates two tabs: `Items` (the catalog + what's on the list) and `Log` (one row per purchase).
3. Deploy > New deployment > Web app: Execute as **Me**, Who has access **Anyone**. Copy the `/exec` URL.

Stores are listed in `STORES` in both `Code.gs` and `index.html`; edit both to add or rename one.

## 2. Host the app
Serve `index.html manifest.json sw.js icon.svg` from GitHub Pages. Local test: `python3 -m http.server 8000`.

## 3. Use
Open the app, tap the gear, paste the URL + token. Install via "Add to Home Screen".

- **List**: items on the list for the store tab you picked. Tick one when it goes in the cart - it is logged as a purchase
  (undo for 6 seconds) and leaves the list. The bar at the bottom is the estimated bill for what's still on the list.
- **Due**: staples you usually rebuy every N days and haven't bought for that long. Tap "+ list" to add.
- **Items**: the whole catalog for the store. Tap "+ list" to put an item on the list. Tap an item's name to edit price, qty, type, notes.
- New items added with the form go straight onto the list.

The Log drives `last_bought`, `times_bought` and `avg_days`. If you edit or delete Log rows by hand, use the sheet menu
**Buy list > Recalculate frequency stats**.
