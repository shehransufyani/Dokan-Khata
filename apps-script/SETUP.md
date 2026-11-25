# Put Store Flow on Google Apps Script

About 10 minutes, once. You need a Google account.

1. Open https://sheets.new and name the sheet **Store Flow**. All shop data, access codes and payments will be stored here.
2. In the sheet, open **Extensions > Apps Script**.
3. In the editor, click `Code.gs`, delete what is there, and paste the whole of
   https://raw.githubusercontent.com/shehransufyani/StoreFlow/main/apps-script/Code.gs
4. Click **+** next to Files, choose **HTML**, name it `Index` (no `.html`), delete what is there, and paste the whole of
   https://raw.githubusercontent.com/shehransufyani/StoreFlow/main/apps-script/Index.html
5. Press **Save** (the disk icon).
6. In the function list at the top choose `setup`, press **Run**, then **Review permissions**. Pick your account. Google says the app is not verified because you made it yourself: press **Advanced**, then **Go to Store Flow (unsafe)**, then **Allow**.
7. Press **Deploy > New deployment**. Click the gear and choose **Web app**. Set **Execute as: Me** and **Who has access: Anyone**. Press **Deploy** and copy the **Web app URL**. That is your app's address.
8. Go back to the sheet and reload it. A **Store Flow** menu appears. Use **Create my master code** to get your own Pro link.

## Running the business from the sheet
- **Payments** sheet: every plan payment customers send from the app lands here, and you get an email. Check the Transaction ID in your bKash, Nagad or Rocket app, click the row, then **Store Flow > Approve selected payment**. It creates a code and shows the link to send on WhatsApp.
- **Codes** sheet: every access code. Change **Expires** to extend a plan, or set **Status** to `disabled` to switch a code off.
- **Data** sheet: each shop's products, sales and expenses, saved automatically. Do not edit by hand.
- **Create access code** in the menu makes a code without a payment (for example a gift or a test).

## Updating the app later
Paste the new `Code.gs` and `Index.html`, then **Deploy > Manage deployments**, edit the web app, choose **New version** and press **Deploy**. The URL stays the same.

## Things that work differently inside Apps Script
- The phone camera scanner does not work there (Google blocks the camera inside Apps Script pages). USB and Bluetooth barcode scanners still work.
- It cannot be installed as an app or used offline. Shops need internet. Unsaved changes wait and upload when the internet returns.
- Old signed codes still open the app, but only codes made in the sheet save data online. Make your master code from the menu.
