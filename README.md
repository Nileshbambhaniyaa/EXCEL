# EXCEL
NEO Store 2025
# NEO Store Annual Sales Analysis (2025)

An end‑to‑end Excel project that cleans, processes, analyzes, and visualizes retail sales data to produce an interactive dashboard with slicers.

---

## 🎯 Project Objectives

* Build a reproducible, resume‑ready Excel analytics project.
* Analyze NEO Store’s 2025 online sales to understand existing customers and plan growth for 2026.
* Deliver an interactive dashboard that auto‑updates when data changes.

---

## 📦 Dataset

Retail transaction data for **NEO Store** (2025). Typical columns:

* `Index` – serial number (ignored in analysis)
* `Order ID` – unique numeric ID (no nulls/duplicates expected)
* `Customer ID` – numeric (no nulls)
* `Gender` – standardized to **Male** / **Women** (normalized from M/W, etc.)
* `Date` – order date (full year, Jan–Dec 2022)
* `Status` – Delivered / Cancelled / Returned / Refunded
* `Channel` – Amazon, Flipkart, Myntra, etc.
* `Category` – product category
* `Size` – normalized to numeric where applicable
* `Quantity` – numeric
* `Currency` – INR
* `Amount` – numeric sales amount
* `City`, `State`, `Country` – location fields (India)

> **Derived columns (created during processing)**
>
> * `Age Group` (bucketed): **Teenager** (<30), **Adult** (30–49), **Senior** (≥50)
> * `Month` (text): `TEXT(Date, "mmm")` (e.g., Jan, Feb, ...)

---

## 🧹 Data Cleaning (Excel)

1. **Standardize categorical labels**

   * Normalize `Gender` values: replace `M`→`Male`, `W`/`Women`→`Women`.
2. **Type checks**

   * Ensure `Order ID`, `Customer ID`, `Quantity`, `Amount` are numeric; no blanks.
3. **Date span sanity check**

   * Confirm inclusive coverage of Jan–Dec 2022; remove outliers if any.
4. **Whitespace / spelling**

   * Fix inconsistent `Category`, `State`, `Channel` spellings.

---

## 🔧 Data Processing (Excel)

* **Age bucketing** (in a new column, then paste values):

  ```excel
  =IF([@Age] >= 50, "Senior", IF([@Age] >= 30, "Adult", "Teenager"))
  ```
* **Month extraction** for trend charts:

  ```excel
  =TEXT([@Date], "mmm")
  ```
* **Number formatting** for axis labels in Millions:

  * Format Code: `0.00,,"M"`

> After computing, **Paste as Values** to avoid workbook slowness.

---

## 📊 Analysis Questions Answered

1. Compare **Sales vs. Orders** in a single chart (monthly).
2. Which **month** had the **highest sales** and the **most orders**?
3. Who purchased more overall: **Men or Women**?
4. What’s the distribution of **Order Status** (Delivered / Returned / Refunded / Cancelled)?
5. Which are the **Top 5 States by Sales**?
6. What’s the relationship between **Age Group × Gender** and **Orders**?
7. Which **Sales Channels** contribute the most orders?

---

## 📈 Charts & How They’re Built

### 1) Sales vs Orders (Combo)

* **PivotTable** rows: `Month`
* Values: `Amount` (Sum), `Order ID` (Count)
* **Chart**: Combo → Column for Sales, **Line** for Orders with **Secondary Axis**
* Axis number format: `0.00,,"M"`

### 2) Men vs Women (Share)

* Values: `Amount` (Sum)
* Columns: `Gender`
* **Chart**: Pie with **Data Labels → Percentage** and **Data Callouts**

### 3) Order Status Mix

* Rows: `Status`
* Values: `Order ID` (Count)
* **Chart**: Pie; rotate first slice to declutter, enable labels

### 4) Top 5 States by Sales

* Rows: `State`
* Values: `Amount` (Sum)
* **Filter → Top 10… → Top 5 by Sum of Amount**
* **Chart**: Horizontal Bar; show **Data Labels**; axis format in Millions

### 5) Age Group × Gender vs Orders

* Rows: `Age Group`
* Columns: `Gender`
* Values: `Order ID` (Count) → **Show Values As → % of Grand Total**
* **Chart**: Clustered Column (stack or cluster as preferred)

### 6) Channel Contribution

* Rows: `Channel`
* Values: `Order ID` (Count) → **% of Grand Total**
* **Chart**: Pie with labels

---

## 🧭 Interactive Dashboard (Slicers)

Add slicers so all charts react to filters.

1. Click any **PivotTable** → **Insert Slicer** → choose: `Month`, `Channel`, `Category`.
2. For each slicer: **Report Connections** → tick **all** PivotTables to link.
3. Arrange slicers on the left; place charts on the canvas. Add a title like:

   > **NEO Store — Annual Report (2025)**
4. Use the **Clear Filter** (red cross) on each slicer to reset the view.

---

## 🔍 Key Insights (from the sample walkthrough)

* **March** shows the **highest sales** and **order volume**.
* **Women** drive the **majority of sales** overall.
* **Delivered** is the dominant order status (healthy fulfillment).
* **Top 5 States** by revenue: **Maharashtra, Karnataka, Uttar Pradesh, Telangana, Tamil Nadu**.
* **Adults (30–49)**, especially **Women**, place the largest share of orders.
* **Amazon** contributes the largest order share (~**35%**), followed by **Flipkart** and **Myntra**.

> **Recommendation:** Target **Women aged 30–49** in **Maharashtra, Karnataka, and Uttar Pradesh** with coupons/offers on **Amazon/Flipkart** to lift conversions.

---
## 🚀 Getting Started

1. **Requirements**: Microsoft Excel 2016 or later (Windows/macOS).
2. **Open** `workbook/NEO_Store_Dashboard.xlsx` (or create your own following steps above).
3. If you change/add data in `data/NEO_store_2025.xlsx`, refresh:

   * **PivotTable Analyze → Refresh All** (updates all charts automatically).

---

## 📝 Notes & Tips

* Prefer **PivotCharts** (not regular charts) if you want slicers to control them.
* After heavy formula work, **Paste as Values** to keep the file fast.
* Use **Design → Grand Totals: Off** to reduce clutter where totals are not needed.
* Keep a consistent color theme and clear titles for readability.

👨‍💻 Author

Nileshkumar Bambhaniya
