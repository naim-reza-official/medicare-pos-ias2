# MediCare POS & IAS 2 FIFO Ledger Suite

A complete, lightweight, serverless **Pharmacy Point of Sale (POS), Inventory Management, and Accounting Ledger System** built entirely on **Google Apps Script** and **Google Sheets**. 

Designed specifically for pharmacies and retail medical stores, this application tracks base inventory in single units (Pcs/Bottles) with automatic Strip and Box conversions, enforces **IAS 2 (International Accounting Standard 2)** FIFO cost valuation, and separates retail markup profit from supplier trade commissions.

---

## Key Features

### 1. Fast Retail Point of Sale (POS)
* **Multi-Unit Selling:** Sell medicines by **Pcs/Bottle**, **Strip (Pata)**, or **Box** on the same invoice with real-time stock deduction in base units.
* **Split-Line Billing:** Add multiple unit lines for the same medicine (e.g., 1 Box + 3 Loose Pcs) in a single voucher.
* **Smart Search & Filtering:** Instant search by brand name, generic name, or manufacturer, plus one-click generic filtering.
* **Customer Auto-Match:** Automatically recognizes returning customers by phone number and pre-fills their details.
* **Cash & Change Calculator:** Built-in cash received and return/change calculation with 80mm thermal receipt printing.
* **Invoice Editing & Full Returns:** Edit past invoices or process full customer returns with automatic FIFO batch restoration.

### 2. IAS 2 Compliant FIFO Inventory & Accounting
* **True FIFO Costing:** Automatically deducts stock from the oldest purchase batches first (`First-In, First-Out`).
* **Commission Accounting (IAS 2):** Supplier trade commissions reduce the net carrying cost of inventory rather than inflating immediate income. Commission profit is recognized only when the stock is actually sold.
* **Real-Time KPI Dashboard:**
  * **Retail Markup Profit:** Selling price minus gross cost.
  * **Realized Commission:** Earned supplier commission on sold items.
  * **Net Realized Profit:** Total gross profit minus inventory write-off losses.
  * **Inventory Write-Off Loss:** Tracks net cost of destroyed/expired stock.

### 3. Wholesale Stock Purchase & Catalog Management
* **Flexible Purchasing:** Record wholesale purchases in Boxes, Strips, or Pcs/Bottles with automatic unit cost breakdown.
* **Master Medicine Catalog:** Configure pack sizes (`Pcs per Strip` and `Strips per Box`), selling prices, compressed product thumbnails, and minimum stock alert thresholds in **Pcs, Strip, or Box**.
* **Selling Price Manager:** Update base unit prices from a dedicated view with automatic Strip and Box price recalculation.

### 4. Expiry Tracker & Automated Low-Stock Email Alerts
* **Batch & Expiry Monitoring:** Highlights expiring and expired batches in both the tracker and the POS screen to prevent accidental dispensing.
* **Batch Disposition:** Return near-expiry batches to suppliers (`0.00 TK` loss) or destroy expired stock (records an IAS 2 inventory write-off).
* **Automated Email Alerts:** Sends automatic HTML email alerts when a sale drops any medicine's stock at or below its configured minimum threshold.

### 5. Audit-Ready Accounting Ledger & CSV Export
* Automatically logs every Retail Sale, Sale Return, Wholesale Purchase, Supplier Return, and Stock Write-Off.
* Filter transactions by date range and entry type, and export clean **CSV statements** for accounting or tax audits.

---

## Tech Stack
* **Backend:** Google Apps Script (`Code.gs`)
* **Database:** Google Sheets (Auto-initialized with structured tabs)
* **Frontend:** Vanilla HTML5, CSS3, and JavaScript (`index.html`) with local caching for fast load times

---

## Installation & Setup Guide

Follow these steps to deploy your own free copy in under 5 minutes:

### Step 1: Create a Google Sheet
1. Go to [Google Sheets](https://sheets.google.com) and create a new blank spreadsheet.
2. Rename the spreadsheet to **MediCare POS Database** (or any name you prefer).

### Step 2: Open Google Apps Script
1. In your Google Sheet, click **Extensions** > **Apps Script** from the top menu.
2. Delete any existing code in the default `Code.gs` file.
3. Copy the entire contents of `Code.gs` from this repository and paste it into the editor.
4. Click **File > Save** (or press `Ctrl + S`).

### Step 3: Add the Frontend Interface
1. In the Apps Script editor, click the **`+`** icon next to **Files** on the left sidebar and select **HTML**.
2. Name the file **`index`** (it will automatically save as `index.html`).
3. Delete the placeholder HTML code, paste the entire contents of `index.html` from this repository, and save (`Ctrl + S`).

### Step 4: Deploy as a Web App
1. Click the blue **Deploy** button at the top right and select **New deployment**.
2. Click the gear icon next to **Select type** and choose **Web app**.
3. Configure the deployment settings:
   * **Description:** `v1.0`
   * **Execute as:** `Me (your email)`
   * **Who has access:** `Only myself` (or `Anyone with Google account` if sharing with staff)
4. Click **Deploy**.
5. Click **Authorize access**, choose your Google account, click **Advanced** > **Go to [Project Name] (unsafe)**, and click **Allow**.
6. Copy the **Web app URL** and bookmark it in your browser.

### Step 5: Initial Configuration
1. Open the Web App URL and navigate to **Shop Settings** in the left sidebar.
2. Enter your Pharmacy Name, Address, Phone Number, Receipt Footer, and Low Stock Alert Email, then click **Save Settings**.
3. Go to **Company & Medicine** to add your medicines, then record your opening stock under **Stock Purchase**.

---

## License
This project is open-source and free to use for personal and commercial pharmacy operations.
