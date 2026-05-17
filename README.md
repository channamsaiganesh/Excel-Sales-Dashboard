# Excel-Sales-Dashboard
# 📊 Excel Sales Dashboard Project

## 📌 Project Overview

This project involves analyzing a sales dataset using Microsoft Excel to create an interactive dashboard and generate business insights.

The workbook contains two sheets:

1. **Data** – Raw transactional data.
2. **Questions and KPIs** – Business questions, required transformations, and key performance indicators to build the report.

The goal is to clean the data, create calculated columns, compute KPIs, and answer business questions through charts, Pivot Tables, and dashboards.

---

## 📂 Workbook Structure

### 📄 Sheet 1: Data

Contains the raw dataset with order and customer information.

Typical columns include:

* Order Date
* Delivery Date
* Customer Gender
* Country
* Product Name
* Quantity
* Amount
* Rating
* Order Channel

### 📄 Sheet 2: Questions and KPIs

Contains:

* Required data transformations
* KPIs to calculate
* Business questions to answer
* Additional insight generation task

---

# 🛠️ Data Transformations

The following transformations are performed before analysis:

### 1. Delivery Duration

Create a column to calculate the number of days between the order date and delivery date.

```excel
=Delivery Date - Order Date
```

### 2. Handle Missing Gender Values

Replace blank values in the Gender column with `Unknown`.

```excel
=IF(TRIM([@Gender])="","Unknown",[@Gender])
```

### 3. Clean Country Names

Remove the word `country` and extra spaces from the Country column.

```excel
=TRIM(SUBSTITUTE([@Country],"country",""))
```

---

# 📈 Key Performance Indicators (KPIs)

The dashboard includes the following KPIs:

| KPI                     | Description               |
| ----------------------- | ------------------------- |
| Total Orders            | Total number of orders    |
| Total Quantity          | Total units sold          |
| Total Amount            | Total revenue generated   |
| Average Rating          | Average customer rating   |
| Average Days to Deliver | Average delivery duration |

---

# ❓ Business Questions Answered

1. **Trend in the Last 13 Weeks**
   Analyze weekly sales performance over the most recent 13 weeks.

2. **How Our Customers Like to Buy**
   Identify preferred purchase channels.

3. **How Many They Buy**
   Analyze quantity purchased per order.

4. **Which Products Are Popular?**
   Determine top-selling products.

5. **Overall Gender Split**
   Understand customer distribution by gender.

6. **Where Do Our Customers Live?**
   Analyze customer geographic distribution.

7. **How Long We Take to Ship the Orders**
   Evaluate delivery performance.

8. **How Happy Are Our Customers?**
   Measure customer satisfaction through ratings.

---

# 🔍 Additional Insights Generated

The project also includes 10 extra insights, such as:

1. Top 5 products by revenue
2. Highest revenue-generating country
3. Monthly sales trend
4. Average rating by product
5. Average delivery time by country
6. Revenue contribution by gender
7. Best-performing order channel
8. Quantity sold by product category
9. Lowest-rated products
10. Weekly order growth rate

---

# 📊 Dashboard Components

The Excel dashboard includes:

* KPI Cards
* Weekly Trend Chart
* Product Popularity Chart
* Channel Distribution Chart
* Gender Split Pie Chart
* Country Analysis Map/Bar Chart
* Delivery Performance Chart
* Customer Rating Visualization
* Slicers for dynamic filtering

---

# 🧰 Excel Features Used

* Pivot Tables
* Pivot Charts
* Slicers
* Calculated Columns
* Excel Functions:

  * `IF`
  * `TRIM`
  * `SUBSTITUTE`
  * `AVERAGE`
  * `COUNT`
  * `SUM`
  * `WEEKNUM`

---

# 📁 Deliverables

* Cleaned dataset with transformation columns
* KPI summary section
* Interactive Excel dashboard
* Insights report

---

# 🚀 How to Use

1. Open `Excel Project.xlsx`.
2. Review the `Data` sheet.
3. Perform the required transformations.
4. Build Pivot Tables and charts.
5. Create dashboard visuals.
6. Use slicers to interact with the report.

---

# 📌 Skills Demonstrated

* Data Cleaning
* Data Analysis
* Excel Functions
* Pivot Tables & Charts
* Dashboard Design
* Business Insight Generation

---

# 🏷️ Project Tags

`Excel` `Data Analysis` `Dashboard` `Business Intelligence` `Pivot Tables` `KPI Reporting`

---

# ⭐ Project Summary

This project demonstrates end-to-end Excel-based business analysis, including data cleaning, KPI creation, dashboard development, and insight generation. It highlights practical Excel skills commonly used in data analyst roles.
