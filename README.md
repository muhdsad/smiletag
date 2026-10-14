# 🏷️ Smile PriceTag Master (Enterprise Retail Price Tag Generator)

A dedicated, fast, and modern software built specifically for supermarkets and retail stores to design, manage, and print high-impact promotional price tags matching the **Smile Supermarket** standard (Yellow MRP Strikethrough Badge, Red Store Savings Badge, ultra-bold black outlined typography, and product photography).

---

## 🌟 Key Features

1. **A4 Sheet Printing (Portrait & Landscape Orientations)**:
   - **🌟 Landscape Sheets (Horizontal)**:
     - **4 Columns × 3 Rows = 12 Tags per A4 Landscape sheet** (spacious, high-impact tags).
     - **4 Columns × 4 Rows = 16 Tags per A4 Landscape sheet** (standard horizontal retail tag format).
   - **📄 Portrait Sheets (Vertical)**:
     - **4 Columns × 5 Rows = 20 Tags per A4 sheet** (standard supermarket format).
     - **4 Columns × 4 Rows = 16 Tags per A4 sheet**.
     - **3 Columns × 4 Rows = 12 Tags per A4 sheet**.
   - Dynamic multi-page support (automatically paginates across multiple A4 sheets).
   - Automatic print orientation detection via `@page { size: A4 landscape/portrait; }`.
   - Dotted cutting guidelines for easy trimming with scissors or paper cutters.
   - Dedicated `@media print` CSS ensuring razor-sharp vector printouts with exact colors and zero browser clipping.

2. **Multi-Source Spreadsheet Import**:
   - **Excel Files (`.xlsx`, `.xls`)**: Upload your product spreadsheet directly.
   - **CSV / TSV Files (`.csv`, `.tsv`)**: Seamless auto-detecting delimited file support.
   - **Google Sheets**:
     - Paste a Google Sheet URL for instant import.
     - Or copy cells from Google Sheets / Excel and paste them directly into the Quick-Paste tab.
   - Smart header mapping automatically detects columns: `Item Name`, `MRP`, `Offer Price`, `Unit`, `Store Name`, `Image`.

3. **Intelligent Image Automation**:
   - **📁 Auto-Detect Images from Folder**: Click the button, select a local folder on your computer containing product photos, and the software automatically matches images to product names using fuzzy search (e.g., `OFFER LITCHI` matches `litchi.jpg`, `apple.png`, etc.).
   - **☁️ Google Drive Link Resolver**: Automatically converts Google Drive share links (`drive.google.com/file/d/...`) into direct high-speed image streams.
   - **Manual Swap & Crop**: Drag-and-drop or click to change any product's photo with zoom and pan controls.

4. **Post-Upload In-Place Editor**:
   - Spreadsheet-style table view where you can modify product titles, MRP, offer prices, and units anytime.
   - Live auto-calculation of Savings (`Save = MRP - Offer Price`) and discounts.
   - Add new tags, duplicate items, reorder, or delete items on the fly.
   - Click directly on any tag in the A4 sheet preview to quick-edit it.

5. **Tag Designer & Positioning Studio**:
   - Fine-tune individual tag visuals: Pan X/Y, Zoom scale, title font size, price font size, and white stroke thickness.
   - High-DPI 300 DPI Canvas exporter: Download individual tags as crisp PNG images.

---

## 🚀 How to Run the Software

### Option A: One-Click Launcher (Linux / Mac)
Run the launcher script in terminal:
```bash
./run.sh
```
This automatically starts the local server and opens the application in your default web browser at `http://localhost:8080`.

### Option B: Windows
Double-click `run.bat`.

### Option C: Direct Browser Launch (No Server Required)
Simply double-click `index.html` or open it in Google Chrome, Microsoft Edge, Firefox, or Safari! All core features (including Excel parsing, CSV import, image folder scanning, and A4 printing) work 100% offline directly in your browser.

---

## 📁 Sample Files Included
- `assets/sample_price_tags.csv` - Ready-to-test CSV file with 16 supermarket items.
- `assets/sample_price_tags.xlsx` - Ready-to-test Excel workbook.
- `assets/sample_images/` - Pre-generated sample produce images.
- `assets/litchi.jpg` - Original reference product photo.

---

## 🖨️ Recommended Print Settings
When clicking **Print A4 Sheet** (or pressing `Ctrl + P`):
1. **Destination**: Select your printer or "Save as PDF".
2. **Layout**: Portrait.
3. **Paper Size**: A4.
4. **Margins**: None (or Default).
5. **Options**: Check **"Background graphics"** so that the Yellow MRP badge, Red Save badge, and images print in full vibrant color!
