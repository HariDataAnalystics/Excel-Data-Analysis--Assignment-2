# Excel Data Cleaning and Transformation

## 📌 Project Overview

This project focuses on **data cleaning and preparation using Microsoft Excel**. As part of my Data Analytics learning journey, I worked with a Product Dataset containing product details such as Product ID, Product Name, Brand Name, Quantity, Category, and Price.

The objective of this assignment was to identify and handle common data quality issues such as **missing values, inconsistent text formats, spelling errors, duplicate records, and unstructured data** before analysis.

This project demonstrates my practical understanding of **Excel data cleaning, transformation, formatting, and conditional formatting techniques**.

---

## 🎯 Objectives

- Identify and handle missing values.
- Standardize inconsistent text formats.
- Correct category spelling errors.
- Identify and remove duplicate records.
- Split and restructure Product ID information.
- Merge columns to create meaningful fields.
- Apply appropriate number and date formatting.
- Use conditional formatting to improve data readability.
- Prepare a clean and analysis-ready dataset.

---

## 📊 Dataset Description

The dataset contains information about different products.

| Column | Description |
|---|---|
| Product ID | Contains product date information and country code |
| Product Name | Name of the product |
| Brand Name | Brand associated with the product |
| Price | Price of the product |
| Quantity | Available quantity |
| Category | Product category |

---

## 🧹 Data Cleaning & Transformation

### 1. Handling Missing Values

#### Price

Missing values were identified in the **Price** column.

For products with missing price information, the missing values were handled using **average price imputation**.

**Excel Formula:**

```excel
=IF(D2="",AVERAGE($D$2:$D$35),D2)
```

This replaces a blank price with the average price while keeping existing prices unchanged.

#### Category

Missing categories were identified and handled using a suitable category-imputation strategy based on the available product information.

The objective was to ensure that the dataset contains meaningful category values before analysis.

---

### 2. Correcting Inconsistent Data

#### Product Name

The **Product Name** column contained inconsistent capitalization and formatting.

The text was standardized using Excel functions such as:

```excel
=PROPER(TRIM(CLEAN(B2)))
```

This helps to:

- Remove unnecessary spaces.
- Remove non-printing characters.
- Standardize capitalization.

#### Category

Typographical errors were identified in the **Category** column.

For example:

```text
Electroni → Electronics
```

The incorrect category values were corrected using **Find & Replace** / Excel text transformation.

Example formula used:

```excel
=IF(F2="","",SUBSTITUTE(F2,"Electroni","Electronics"))
```

---

## 🗑️ 3. Removing Duplicate Records

The dataset was checked for duplicate rows based on the **complete row contents**.

Excel's **Remove Duplicates** feature was used to identify and remove duplicate records where applicable.

This helps ensure that each product record is represented only once in the cleaned dataset.

---

## 🔄 4. Splitting and Merging Data

### Splitting Product ID

The Product ID follows a structure similar to:

```text
28-JAN-US
```

The Product ID was separated into:

- **Manufacturing Date**
- **Country Code**

The country code was extracted using:

```excel
=RIGHT(A2,2)
```

The manufacturing date was extracted and converted into a proper Excel date format.

Example:

```excel
=DATE(2026,MONTH(DATEVALUE("1-"&MID(A2,4,3))),VALUE(LEFT(A2,2)))
```

### Merging Brand Name and Product Name

The **Brand Name** and **Product Name** fields were combined to create a new column:

```text
Product Brand
```

Example:

```text
Dell + Laptop → Dell Laptop
```

This creates a more descriptive field for analysis.

---

## 💰 5. Number Formatting

### Price

The **Price** column was formatted using a **Currency format** to improve readability and clearly represent monetary values.

Example:

```text
1000 → $1,000.00
```

### Manufacturing Date

The Manufacturing Date column was formatted as:

```text
DD-MM-YYYY
```

Example:

```text
28-01-2026
```

---

## 🎨 6. Conditional Formatting

### Price

Conditional formatting was applied to the **Price** column using:

- Data Bars
- Color Scales

This makes it easier to visually identify relatively high and low product prices.

### Category

A custom conditional formatting rule was created to highlight products belonging to:

```text
Electronics
```

This allows specific product categories to be identified quickly within the dataset.

---

## 🛠️ Tools & Skills Used

### Tools

- Microsoft Excel
- GitHub

### Excel Skills

- Data Cleaning
- Missing Value Handling
- Average Imputation
- Find & Replace
- Text Standardization
- `IF()`
- `AVERAGE()`
- `PROPER()`
- `TRIM()`
- `CLEAN()`
- `SUBSTITUTE()`
- `RIGHT()`
- `LEFT()`
- `MID()`
- `DATE()`
- `DATEVALUE()`
- Duplicate Removal
- Data Formatting
- Date Formatting
- Currency Formatting
- Conditional Formatting
- Data Bars
- Color Scales

---

## 📁 Project Files

```text
Excel-Data-Cleaning-Transformation/
│
├── Dataset/
│   └── Product Dataset.xlsx
│
├── Assignment/
│   └── Assignment 2 - Data Cleaning and Transformation.pdf
│
└── README.md
```

---

## 🔄 Data Cleaning Workflow

```text
Raw Product Dataset
        ↓
Check Missing Values
        ↓
Handle Missing Price / Category
        ↓
Standardize Product Names
        ↓
Correct Category Typos
        ↓
Remove Duplicate Records
        ↓
Split Product ID
        ↓
Merge Brand + Product Name
        ↓
Format Price & Date
        ↓
Apply Conditional Formatting
        ↓
Clean & Analysis-Ready Dataset
```

---

## 📈 Key Learning Outcomes

Through this project, I gained practical experience in:

- Understanding the importance of data quality.
- Cleaning real-world style datasets using Excel.
- Handling missing and inconsistent data.
- Transforming unstructured information into useful columns.
- Using Excel formulas for data preprocessing.
- Applying formatting techniques for better data readability.
- Preparing datasets for further analysis and visualization.

---

## 🚀 About My Data Analytics Journey

I am a **Final-Year Biomedical Engineering student** currently developing my skills in **Data Analytics**.

I am building my portfolio by completing practical projects using:

```text
Excel → SQL → Python → Power BI
```

This project is part of my learning journey toward becoming a **Data Analyst**.

---

## 👤 Author

**Hariharan A**

Aspiring Data Analyst | Final-Year Biomedical Engineering Student

Skills: **Excel | SQL | Python | Power BI**

---

## ⭐ Project Status

**Completed ✅**

More Data Analytics projects will be added to this portfolio as I continue learning and developing my skills.
