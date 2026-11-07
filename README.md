# Hot Diggity Dog website

One-page site for the Hot Diggity Dog food truck (Kasantú, select West African cuisine).

## Files
- `index.html`: the whole website
- `truck.jpg`: truck photo
- `logo.png`: logo
- `img/`: mascot, signature dog photos, and food photos

## Put it online with GitHub Pages
1. Create a new **public** repository, for example `hot-diggity-dog`.
2. Click **Add file > Upload files**, drag in `index.html`, `truck.jpg` and `logo.png`, then click **Commit changes**.
3. Go to **Settings > Pages**. Under "Build and deployment," set Source to **Deploy from a branch**, Branch to **main**, folder to **/ (root)**, then click **Save**.
4. Wait a minute or two and refresh. The site will be at `https://YOUR-USERNAME.github.io/hot-diggity-dog/`.

## Easy edits
Open `index.html` on GitHub, click the pencil icon, and change the **EASY-EDIT SETTINGS** block near the top:
- `kasantu`: the single West African dish card next to "Choose your dog" (leave `name` empty to show "coming soon"; set `show: false` to hide it)
- `schedule`: where the truck will be
- `phone`, `phoneDigits`, `email`
- `ordering.enabled`: set to `false` to hide the order buttons and cart
- `ordering.sheetUrl`: the Google Apps Script Web app URL; orders then go straight into the Google Sheet (see `google-sheet-orders/Code.gs`)
- `ordering.payUrl`: the Cloudflare pay-helper URL (see `square-pay-helper/worker.js`); turns on "Pay now with card" through Square
- `ordering.squareUrl`: paste the Square Online ordering link to send all orders to Square instead of the text-message cart

Menu items and prices are in the lists right below it: `SIGNATURE` (hot dog styles), `CHICKEN` (chicken and sandwiches), `COMBOS`, `MORE` (nachos, fries, parfait) and `SIDES`.
Click **Commit changes** and the live site updates in about a minute.
