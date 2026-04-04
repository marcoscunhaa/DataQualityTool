📊 Data Quality Tool
===============================

Is a Python-based tool for auditing and cleaning datasets, identifying inconsistencies, and preparing data for analysis and visualization.

* * *

📌 Overview
-----------

This tool is divided into three main stages:

1. **Exploratory Analysis (`exploratory.ipynb`)**
2. **Data Quality Audit (`audit.py`)**
3. **Data Cleaning (`cleaning.py`)**

The workflow follows a practical industry approach:

> 🔍 First understand the data → 📊 Identify issues → 🧹 Clean and standardize

* * *

## ⚠️ 1. Exploratory Analysis

The notebook `exploratory.ipynb` is used for an initial inspection of the dataset.

### Key steps:

* Preview data (`head`, `sample`)
* Check structure and data types (`info`)
* Generate descriptive statistics (`describe`)
* Identify duplicates
* Inspect categorical inconsistencies (e.g., Product, City, State, Country)

This step helps to:

* Quickly understand the dataset

* Detect obvious issues before automation

* Guide the audit and cleaning strategy
  
  ## 📉 2.  Data Quality Audit

The script `audit.py` performs a **comprehensive data quality scan** and generates a structured Excel report.

### ✅ Checks performed:

* Missing values
* Mixed data types in columns
* Duplicate rows
* Duplicate IDs
* Price columns stored as text
* Invalid characters in numeric fields
* Outlier detection (IQR method)
* Invalid date formats
* Text inconsistencies (case/spacing)
* Contact fields with incorrect data types
* Negative values in quantity fields
* Potential category abbreviations

* * *

### 📈 Output

An Excel report is generated with:

* **Quality Summary**
  * Total rows, columns, cells
  * Total errors
  * Error rate (%)
  * Data Quality Score (%)
  * Quality classification (Excellent / Good / Regular / Poor)
* **Error Details**
  * Column
  * Error type
  * Count
  * Percentage

* * *

🧹 3. Data Cleaning
-------------------

The script `cleaning.py` applies automated corrections based on common data quality issues.

### 🔧 Cleaning operations:

* Remove duplicate rows (based on ID when available)
* Standardize and parse currency values
* Normalize date formats (converted to US format: `MM/DD/YYYY`)
* Fill missing discount values with `0`
* Replace null text values with `"Unknown"` (non-critical fields)
* Standardize text formatting (case normalization)
* Convert contact fields to string
* Normalize category names using mapping
* Remove invalid or duplicate records

* * *

### 📤 Output

* **Cleaned Data** → ready for analysis
* **Discarded Data** → removed records for traceability

Both are exported to Excel.

* * *

⚙️ Technologies Used
--------------------

* Python
* Pandas
* NumPy
* OpenPyXL

* * *

🚀 How to Use
-------------

### 1. Exploratory Analysis

**Run the notebook:**

exploratory.ipynb

### 2. Run Data Audit

```
python audit.py
```

### 3. Run Data Cleaning

```
python cleaning.py
```

* * *

📁 Input Data
-------------

The project expects a dataset such as:

Walmart Inventory.csv

Supports:

* `.csv`
* `.txt`
* `.xlsx`
* `.xls`

* * *

💡 Use Cases
------------

This tool can be applied to:

* Business datasets (sales, inventory, CRM)
* Data preparation for dashboards (Power BI, Looker Studio)
* Data pipelines (ETL processes)
* Freelance data cleaning projects (Upwork, Fiverr)

* * *

🎯 Key Highlights
-----------------

* End-to-end data quality workflow

* Automated auditing with scoring system

* Real-world cleaning logic (currency, dates, categories)

* Excel reporting for business-friendly output

* Modular and reusable design

*  
