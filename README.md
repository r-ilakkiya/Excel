# Excel
Excel Practice Files - Assignments and Projects
# 📊 Product Data Exploration Using Microsoft Excel

## 📌 1. Project Overview & Purpose

This project focuses on **Product Data Exploration using Microsoft Excel**.

The dataset contains information about different products, including product names, brands, prices, quantities, categories, price ranges, and country codes.

The main objective of this project is to understand and analyse product information using Excel functions and formulas.

This project demonstrates basic Excel Data Analytics skills, including data exploration, calculations, logical functions, conditional functions, and text manipulation.

### 🎯 Project Objectives

- Explore and understand the product dataset.
- Calculate total, average, minimum, and maximum product prices.
- Categorize products based on their prices using the IF function.
- Analyse product information using SUMIF and COUNTIF.
- Extract information from Product IDs using LEFT, RIGHT, and MID functions.
- Practise basic data analysis techniques using Microsoft Excel.

## 📋 2. Dataset Overview

| Attribute | Details |
|---|---|
| Dataset Name | Product Data Exploration |
| Total Records | 34 |
| Total Columns | 10 |
| File Format | Excel (.xlsx) |
| File Size | Approximately 16.6 KB |
| Missing Values | 0 |
| Worksheets | Dataset, Task |
| Tool Used | Microsoft Excel |

## 🗂️ 3. Dataset Structure

The dataset contains the following 10 columns:

| Column Name | Description |
|---|---|
| Product ID | Unique product identifier containing day, month, and country code |
| Product Name | Name of the product |
| Brand Name | Brand associated with the product |
| Price ($) | Individual product price in US dollars |
| Quantity | Quantity associated with the product |
| Category | Classification of products, such as Electronics, Fashion, and Kitchen |
| Price Range | Product price classification based on price |
| Day | Day extracted from Product ID |
| Country Code | Country code extracted from Product ID |
| Month | Month extracted from Product ID |

## 🛠️ 4. Tools & Excel Functions Used

**Primary Tool: Microsoft Excel**

The following Excel functions were used to explore and analyse the dataset.

### 🔹 1. SUM Function

**Purpose:** Calculates the total price of all products.

```excel
=SUM(D2:D35)
```

**Result:** $10,100

### 🔹 2. COUNT Function

**Purpose:** Counts the number of numeric product prices.

```excel
=COUNT(D2:D35)
```

**Result:** 34

### 🔹 3. AVERAGE Function

**Purpose:** Calculates the average product price.

```excel
=AVERAGE(D2:D35)
```

**Result:** $297.06

### 🔹 4. MIN Function

**Purpose:** Identifies the lowest product price.

```excel
=MIN(D2:D35)
```

**Result:** $30

### 🔹 5. MAX Function

**Purpose:** Identifies the highest product price.

```excel
=MAX(D2:D35)
```

**Result:** $1,000

### 🔹 6. IF Function

**Purpose:** Categorizes products based on their prices.

**Conditions:**
- If Price is greater than or equal to $500 → High Price
- If Price is less than $500 → Standard Price

```excel
=IF(D2>=500,"High Price","Standard Price")
```

**Result:** Created a new column named `Price Range`.

### 🔹 7. SUMIF Function

**Purpose:** Calculates the total price of products belonging to the Electronics category.

```excel
=SUMIF(F2:F35,"Electronics",D2:D35)
```

**Result:** $8,050

### 🔹 8. COUNTIF Function

**Purpose:** Counts the number of products with a price below $100.

```excel
=COUNTIF(D2:D35,"<100")
```

**Result:** 11 products

### 🔹 9. LEFT Function

**Purpose:** Extracts the first two characters from Product ID to identify the day.

```excel
=LEFT(A2,2)
```

**Example:** 28-JAN-US → 28

### 🔹 10. RIGHT Function

**Purpose:** Extracts the last two characters from Product ID to identify the country code.

```excel
=RIGHT(A2,2)
```

**Example:** 28-JAN-US → US

### 🔹 11. MID Function

**Purpose:** Extracts the month from Product ID.

```excel
=MID(A2,4,3)
```

**Example:** 28-JAN-US → JAN

## 📈 5. Key Analysis Results

The following results were obtained from the dataset:

| Analysis | Result |
|---|---:|
| Total Product Price | $10,100 |
| Total Product Records | 34 |
| Average Product Price | $297.06 |
| Minimum Product Price | $30 |
| Maximum Product Price | $1,000 |
| Electronics Category Total Price | $8,050 |
| Products Below $100 | 11 |

### 📊 Product Category Distribution

| Category | Number of Products |
|---|---:|
| Electronics | 15 |
| Fashion | 7 |
| Kitchen | 6 |
| Outdoor | 3 |
| Accessories | 3 |

### 💡 Key Insights

- Electronics is the largest category, with 15 product records.
- Fashion has 7 product records.
- Kitchen contains 6 product records.
- The lowest product price is $30.
- The highest product price is $1,000.
- 8 products are classified as High Price.
- 26 products are classified as Standard Price.

## 📁 6. Repository Structure

```text
Product-Data-Exploration/
│
├── 📊 Data Exploration Assignment 1.xlsx
│
└── 📄 README.md
```

| File | Description |
|---|---|
| Data Exploration Assignment 1.xlsx | Original Excel workbook containing the dataset and analysis tasks |
| README.md | Project documentation and analysis explanation |

## 📚 7. Skills Demonstrated

Through this project, I practised the following Excel skills:

- Basic Data Exploration
- Data Analysis
- Mathematical Functions
- Logical Functions
- Conditional Functions
- Text Functions
- Data Classification
- Working with Structured Datasets
- Extracting Meaningful Insights

## 📜 8. Data Source & License

**Data Source:** Excel assignment dataset provided for learning and practice.

**Collection Method:** The original data collection method was not specified.

**License:** No license information was provided with the original dataset. Permission should be confirmed before redistribution or public reuse.

## 👩‍💻 9. Author

**Ilakkiya R.**

🎯 Aspiring Data Analyst

🛠️ **Tool:** Microsoft Excel

📊 **Project:** Product Data Exploration

⭐ Thank you for visiting my project! I'm continuously learning and improving my Data Analytics skills.
