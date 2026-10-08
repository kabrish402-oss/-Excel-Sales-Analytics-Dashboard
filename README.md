# 📊 Excel Sales Analytics Dashboard

An end-to-end **Sales Analytics project built in Microsoft Excel** using a real-world-style sales dataset. The project focuses on data cleaning, Excel formulas, lookups, business analysis, and interactive dashboard creation.
## 🎯 Project Objective

The main goal of this project was to transform raw sales data into meaningful business insights using Excel.

The project demonstrates how Excel can be used to:

* Clean and prepare raw data
* Create calculated columns
* Apply logical formulas
* Perform lookups
* Analyze sales performance
* Compare regions, categories, and sales representatives
* Identify trends
* Build a professional sales dashboard

---

## 📁 Dataset

The dataset contains sales transaction information including:

* Order ID
* Order Date
* Customer Name
* Region
* Product Category
* Quantity
* Unit Price
* Sales Rep

The dataset also contains duplicate Order IDs and inconsistent customer-name capitalization, which were handled during the data-cleaning process.

---

## 🛠️ Excel Skills Demonstrated

### Data Cleaning

* Customer name standardization using `PROPER`
* Duplicate Order ID identification
* Data formatting
* Structured Excel Tables

### Excel Formulas

#### Basic Functions

* `SUM`
* `AVERAGE`
* `MIN`
* `MAX`
* `COUNT`

#### Logical Functions

* `IF`
* `IFS`
* `AND`
* `OR`
* `IFERROR`

#### Lookup Functions

* `XLOOKUP`
* `VLOOKUP`
* `INDEX`
* `MATCH`

### Data Analysis

* Sales by Region
* Sales by Product Category
* Sales by Sales Representative
* Monthly Sales Trend
* Total Sales
* Average Order Value
* Total Quantity
* Minimum and Maximum Sales

### Visualization

* KPI cards
* Bar charts
* Column charts
* Line charts
* Sales dashboard

---

## 📊 Dashboard

The final dashboard provides an overview of sales performance through key metrics and visualizations.

### Key Performance Indicators

* 💰 Total Sales
* 📦 Total Orders
* 📊 Average Order Value
* 🔢 Total Quantity

### Visualizations

* Sales by Region
* Sales by Product Category
* Sales by Sales Representative
* Monthly Sales Trend

---

## 📂 Project Structure

```text
Excel-Sales-Analytics-Dashboard/
│
├── Sales_Analytics_Dashboard.xlsx
├── README.md
│
└── screenshots/
    └── dashboard.png
```

---

## 🔎 Business Questions Answered

This project uses Excel to answer questions such as:

1. What is the total sales revenue?
2. What is the average order value?
3. What is the highest-value order?
4. What is the lowest-value order?
5. How many orders were placed?
6. Which region generated the most sales?
7. Which product category generated the most revenue?
8. Which sales representative performed best?
9. How do sales change over time?
10. Which orders should be considered high-value or priority orders?

---

## 🧮 Example Calculated Columns

### Total Sales

```excel
=Quantity*Unit Price
```

### Order Size

```excel
=IF(Total Sales>=1000,"Large","Small")
```

### Sales Level

```excel
=IFS(
Total Sales>=2000,"Very High",
Total Sales>=1000,"High",
Total Sales>=500,"Medium",
TRUE,"Low"
)
```

### Priority Order

```excel
=IF(AND(Total Sales>=1000,Quantity>=5),"Priority","Normal")
```

### Special Order

```excel
=IF(OR(Quantity>=10,Total Sales>=2000),"Special","Normal")
```

---

## 💡 Key Learning

This project helped me understand the complete Excel analytics workflow:

```text
Raw Data
   ↓
Data Cleaning
   ↓
Calculated Columns
   ↓
Logical & Lookup Formulas
   ↓
Data Analysis
   ↓
Business Insights
   ↓
Data Visualization
   ↓
Dashboard
```

---

## 🚀 Future Improvements

Possible future improvements include:

* Adding interactive slicers
* Adding more advanced PivotTables
* Adding profit and profit-margin analysis
* Adding customer segmentation
* Connecting the dataset to Power BI
* Automating the dashboard with larger datasets
* Building the same analysis using SQL and Python

---

## 🧰 Tools Used

* **Microsoft Excel**
* Excel Tables
* Excel Formulas
* Lookup Functions
* PivotTable-style analysis
* Excel Charts
* Dashboard Design

---

## 👩‍💻 Author

**Abrish Khan**

Computer Science Student | Aspiring Data Analyst

### Current Focus

`Excel` • `SQL` • `Power BI` • `Python` • `Data Analytics`

---

⭐ If you find this project useful, feel free to explore the workbook and the analysis process.
