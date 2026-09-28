# Product-Analysis---Ass2
# Excel Data Cleaning & Transformation

## 📌 Project Overview

This project is part of my **Data Analytics learning journey**, where I am building practical skills in **Microsoft Excel and data preprocessing**.

The objective of this project is to clean, standardize, transform, and format a product dataset so that it can be used reliably for further analysis.

Real-world datasets frequently contain issues such as missing values, inconsistent text, duplicate records, poorly structured columns, and inconsistent formatting. This project demonstrates how these issues can be identified and addressed using Excel's built-in data cleaning and transformation features.

---

## 🎯 Project Objectives

The main objectives of this project were to:

* Identify and handle missing values
* Standardize inconsistent text data
* Correct category names and spelling inconsistencies
* Identify and remove duplicate records
* Split structured information into separate columns
* Merge related columns into a meaningful field
* Apply appropriate number and date formatting
* Use conditional formatting to improve data readability
* Prepare the dataset for further analysis

---

## 📊 Dataset

**Dataset:** Product Dataset

The dataset contains product-related information including:

| Column             | Description                                               |
| ------------------ | --------------------------------------------------------- |
| Product ID         | Unique identifier containing date and country information |
| Manufacturing Date | Date associated with product manufacturing                |
| Country Code       | Country identifier extracted from Product ID              |
| Product Brand      | Combined brand and product information                    |
| Price ($)          | Product price                                             |
| Quantity           | Available/product quantity                                |
| Category           | Product category                                          |

The working dataset contains product information across categories such as **Electronics, Fashion, Kitchen, Accessories, and Outdoor**.

---

## 🛠️ Tools & Technologies

* **Microsoft Excel**
* Excel Formulas & Functions
* Find & Replace
* Remove Duplicates
* Text Transformation
* Conditional Formatting
* Number & Date Formatting
* Data Cleaning & Preparation

---

# 🔍 Data Cleaning & Transformation Process

## 1. Handling Missing Values

The dataset was checked for missing values, particularly in the **Price** and **Category** columns.

### Price

Missing or invalid price values were identified and handled using an appropriate replacement strategy rather than leaving incomplete records in the cleaned dataset.

An Excel formula-based approach was used where required:

```excel
=IF(OR(ISBLANK(D2),D2=0),AVERAGE($D$2:$D$35),D2)
```

This approach checks whether the price is blank or zero and replaces it with an appropriate average value.

### Category

Missing categories should be investigated using related product information such as:

* Product Name
* Brand Name
* Similar products
* Existing category patterns

Where a reliable category cannot be determined, the value should be retained as **Unknown** rather than making an unsupported assumption.

---

## 2. Correcting Inconsistent Data

### Product Name Standardization

Product names were reviewed for:

* Inconsistent capitalization
* Leading/trailing spaces
* Unnecessary characters
* Formatting inconsistencies

Excel text functions such as the following can be used for standardization:

```excel
=PROPER(TRIM(CLEAN(C2)))
```

This helps remove unnecessary spaces and non-printing characters while standardizing capitalization.

### Category Standardization

The **Category** column was reviewed for spelling mistakes and inconsistent category names.

Excel's **Find and Replace** functionality was used to locate incorrect values and replace them with standardized category names.

Example:

```text
Incorrect Category → Correct Category
```

This ensures that the same category is represented consistently throughout the dataset.

---

## 3. Removing Duplicate Records

The dataset was checked for duplicate rows using the complete set of columns.

### Excel Process

**Data → Remove Duplicates**

All relevant columns were selected so that duplicate records could be identified based on the entire row.

Duplicate records were removed to ensure that each product record appears only once in the cleaned dataset.

This improves:

* Data accuracy
* Record uniqueness
* Reliability of future analysis

---

## 4. Splitting the Product ID

The original **Product ID** contains multiple pieces of information.

Example:

```text
28-JAN-US
```

The Product ID was transformed into separate fields:

```text
Manufacturing Date → 28-JAN-2026
Country Code       → US
```

Excel text functions can be used to extract these values.

### Manufacturing Date

```excel
=LEFT(A2,6)
```

The extracted date information was then converted/formatted as a date.

### Country Code

```excel
=RIGHT(A2,2)
```

This extracts the country code from the end of the Product ID.

This transformation makes the information easier to analyze independently.

---

## 5. Merging Brand Name & Product Name

The **Brand Name** and **Product Name** fields were combined into a single field called:

### `Product Brand`

Example:

```text
Brand Name: Dell
Product Name: Laptop

Product Brand: Dell Laptop
```

An Excel concatenation formula can be used:

```excel
=A2&" "&B2
```

This creates a single descriptive field while preserving the relationship between the brand and product.

---

## 6. Number & Date Formatting

### Price Formatting

The **Price** column was formatted as currency to make monetary values easier to read.

Example:

```text
1000 → $1,000.00
```

Currency formatting improves consistency and readability when presenting financial information.

### Manufacturing Date

The **Manufacturing Date** column was formatted using:

```text
DD-MM-YYYY
```

Example:

```text
28-JAN-2026 → 28-01-2026
```

This provides a consistent date representation across the dataset.

---

# 🎨 Conditional Formatting

Conditional formatting was applied to improve the visual interpretation of the cleaned dataset.

## Price Data Bars / Color Scales

A conditional formatting rule was applied to the **Price** column using Excel's:

**Conditional Formatting → Data Bars / Color Scales**

This makes it easier to visually identify relatively lower and higher product prices.

## Electronics Category

A custom conditional formatting rule was created for the **Category** column to highlight records where:

```text
Category = Electronics
```

This provides a quick visual way to identify products belonging to the Electronics category.

---

# 📁 Project Structure

```text
Excel-Data-Cleaning-Transformation/
│
├── Assignment 2 - Data Cleaning and Transformation.xlsx
│
├── README.md
│
└── screenshots/
    ├── missing-values.png
    ├── find-replace.png
    ├── duplicate-removal.png
    ├── column-splitting.png
    ├── column-merging.png
    └── conditional-formatting.png
```

> Screenshot filenames can be adjusted to match the actual files uploaded to the repository.

---

# 📸 Evidence & Documentation

The project includes screenshots demonstrating the major cleaning operations performed in Excel.

### Missing Value Handling

Demonstrates how missing values were identified and handled.

### Find & Replace

Shows the correction of inconsistent product/category text.

### Duplicate Removal

Demonstrates the Excel **Remove Duplicates** operation.

### Splitting & Merging

Shows how Product ID was separated into meaningful fields and how brand/product information was combined.

### Conditional Formatting

Demonstrates visual highlighting applied to prices and Electronics category records.

---

# 📈 Key Learning Outcomes

Through this project, I developed practical experience with:

* Data cleaning in Excel
* Missing-value handling
* Text standardization
* Find & Replace
* Duplicate detection and removal
* Text extraction using Excel formulas
* Column transformation
* Data concatenation
* Currency formatting
* Date formatting
* Conditional formatting
* Preparing datasets for analysis

---

# 💡 Why This Project Matters

Data analysis does not begin with visualization or statistical analysis. **Data quality is an important first step.**

An analyst needs to ensure that the underlying dataset is:

* Accurate
* Consistent
* Complete
* Structured
* Readable
* Ready for analysis

This project demonstrates my understanding of these fundamental data-preparation practices using Microsoft Excel.

---

# 🚀 Future Improvements

As I continue developing my data analytics skills, I plan to extend this project by:

* Performing exploratory data analysis
* Creating Excel PivotTables
* Building interactive dashboards
* Analyzing category-wise sales/product metrics
* Identifying pricing patterns
* Creating visual reports
* Recreating the cleaning workflow using **SQL**
* Recreating the same workflow using **Python and Pandas**
* Comparing Excel, SQL, and Python approaches to data cleaning

---

# 👨‍💻 About Me

I am an aspiring **Data Analyst** currently developing my skills in data cleaning, data analysis, visualization, and business intelligence.

My learning journey focuses on building practical projects and documenting the process so that I can gradually develop a strong data analytics portfolio.

### Current Learning Focus

* 📊 Microsoft Excel
* 🧹 Data Cleaning
* 📈 Data Analysis
* 🗄️ SQL
* 🐍 Python
* 📊 Data Visualization
* 📋 Power BI

---

## ⭐ Portfolio

This repository is part of my ongoing **Data Analytics Portfolio**, where I document projects and practical exercises as I build my skills toward a career in data analytics.
