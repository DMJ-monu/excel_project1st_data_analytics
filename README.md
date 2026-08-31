
# 📊 Project 1 — Excel Fundamentals & Data Analysis

## 🧰 Microsoft Excel Features Used

### 1. Logical Functions

Logical functions were used to apply conditions and make decisions within the dataset.

**Functions used / practiced:**

* `IF()`
* `IFS()`
* `AND()`
* `OR()`
* `NOT()`
* `IFERROR()`

**Used for:**

* Creating conditions
* Categorizing data
* Checking values
* Handling errors
* Creating logical classifications

Example:

```excel
=IF(B2>=50000,"High Salary","Low Salary")
```

---

### 2. Mathematical Functions

Mathematical functions were used for calculations and numerical analysis.

**Functions include:**

* `SUM()`
* `SUMIF()`
* `SUMIFS()`
* `ROUND()`
* `ROUNDUP()`
* `ROUNDDOWN()`
* `ABS()`
* `MIN()`
* `MAX()`

**Used for:**

* Salary calculations
* Totals
* Conditional calculations
* Minimum and maximum values
* Numerical transformations

---

### 3. Statistical Functions

Statistical functions were used to understand salary and job-market data.

**Functions include:**

* `AVERAGE()`
* `MEDIAN()`
* `MODE()`
* `COUNT()`
* `COUNTA()`
* `COUNTIF()`
* `COUNTIFS()`
* `MIN()`
* `MAX()`

**Used for:**

* Average salary
* Median salary
* Job counts
* Skill frequency
* Salary comparison
* Statistical summaries

---

### 4. Data Formatting

Excel formatting tools were used to make the dataset and final analysis easier to understand.

**Formatting techniques:**

* Number formatting
* Currency formatting
* Percentage formatting
* Date formatting
* Font and cell formatting
* Borders
* Alignment
* Cell styles
* Conditional formatting

**Purpose:**

To present data in a clean, readable, and professional format.

---

### 5. Conditional Formatting

Conditional formatting was used to highlight important values and patterns.

Examples:

* Highlighting high salaries
* Highlighting low salaries
* Identifying important values
* Comparing numerical values
* Using color scales
* Using data bars
* Highlighting duplicate or specific values

---

### 6. Charts & Visualization

Charts were created to convert numerical data into visual insights.

**Charts used:**

* Column Chart
* Bar Chart
* Line Chart
* Pie Chart
* Doughnut Chart
* Area Chart

**Used for analyzing:**

* Salary distribution
* Job demand
* Skill popularity
* Job categories
* Comparisons between different values

---

### 7. Data Sorting & Filtering

Excel sorting and filtering tools were used to explore the dataset.

Examples:

* Sorting salaries from highest to lowest
* Filtering specific job roles
* Filtering specific skills
* Filtering locations
* Finding specific categories
* Removing unnecessary records from analysis

---

### 8. Excel Tables

Excel Tables were used to organize the dataset into a structured format.

Benefits include:

* Automatic formatting
* Structured references
* Easier filtering
* Dynamic ranges
* Easier data management

---

### 9. Excel Collaboration

Excel collaboration features were used to understand and work with spreadsheets in a structured way.

This includes:

* Sharing workbooks
* Reviewing data
* Organizing worksheets
* Maintaining consistent formatting
* Working with structured spreadsheet data

---

### 10. Basic Data Analysis

The project used Excel to perform basic exploratory analysis.

The analysis included:

* Finding trends
* Comparing values
* Identifying highest and lowest values
* Calculating statistics
* Comparing categories
* Finding relationships between variables
* Extracting useful insights

---

## 📌 Project 1 Skill Summary

```text
Microsoft Excel
│
├── Logical Functions
├── Mathematical Functions
├── Statistical Functions
├── Data Formatting
├── Conditional Formatting
├── Sorting & Filtering
├── Excel Tables
├── Charts & Visualization
├── Basic Data Analysis
└── Excel Collaboration
```

---

# 📈 Project 2 — Advanced Excel Data Analysis & Business Intelligence

Project 2 uses more advanced Excel capabilities for **data transformation, analysis, data modeling, and dashboard development**.

---

## 🧰 Advanced Excel Features Used

### 1. Pivot Tables

Pivot Tables were used to summarize and analyze large datasets.

They were used for:

* Job-count analysis
* Salary analysis
* Job-role comparison
* Country-wise analysis
* Skill analysis
* Employment-type analysis
* Aggregating large amounts of data

Pivot Tables made it possible to analyze thousands of records quickly without manually calculating every value.

---

### 2. Advanced Pivot Table Analysis

Advanced Pivot Table features were used to create more detailed analysis.

Features include:

* Row and column grouping
* Value field settings
* Sorting
* Filtering
* Calculated values
* Percentage calculations
* Ranking
* Grouping categories
* Drill-down analysis

---

### 3. Pivot Charts

Pivot Charts were created from Pivot Tables to visually represent the analysis.

Used for:

* Job demand comparison
* Salary comparison
* Job-role analysis
* Country analysis
* Employment-type analysis
* Skill distribution

Pivot Charts also make it easier to create dashboards that update when the underlying Pivot Table changes.

---

### 4. Analysis ToolPak

The **Analysis ToolPak** was used for statistical and analytical tasks.

It provides additional Excel analysis capabilities such as:

* Descriptive Statistics
* Regression
* Correlation
* Sampling
* Moving Average
* Histogram
* Other statistical analysis techniques

**Purpose:**

To perform more advanced statistical analysis beyond standard Excel functions.

---

### 5. Data Analysis

Advanced analysis techniques were used to extract meaningful information from the dataset.

The analysis included:

* Descriptive statistics
* Salary analysis
* Job-demand analysis
* Category comparison
* Distribution analysis
* Relationship analysis
* Trend identification

---

# 🔄 Power Query

Power Query was one of the major advanced tools used in Project 2.

It was used for **data extraction, cleaning, transformation, and preparation** before the final analysis.

---

### 6. Power Query — Data Import

Power Query can connect to different data sources and bring the data into Excel in a structured format.

It helps reduce repetitive manual data preparation.

---

### 7. Power Query — Data Cleaning

Power Query was used for data-cleaning operations such as:

* Removing unnecessary columns
* Removing rows
* Renaming columns
* Changing data types
* Removing duplicates
* Handling missing values
* Replacing values
* Splitting columns
* Merging columns
* Filtering records

---

### 8. Advanced Transformation

Power Query's transformation features were used to prepare the dataset for analysis.

Examples include:

* Split Column
* Merge Column
* Replace Values
* Fill Down
* Fill Up
* Group By
* Pivot Column
* Unpivot Column
* Change Data Type
* Conditional Column
* Custom Column

These transformations help convert raw data into an analysis-ready format.

---

### 9. Append Queries

**Append** was used to combine datasets vertically.

Conceptually:

```text
Table A
   +
Table B
   +
Table C
   ↓
Combined Table
```

This is useful when multiple tables contain the same type of information and need to be combined into one dataset.

---

### 10. Merge Queries

**Merge** was used to combine related datasets horizontally based on matching columns.

Conceptually:

```text
Table A ─────┐
             ├── Match Key ──→ Combined Dataset
Table B ─────┘
```

For example, two tables can be connected using a common field such as:

* Job ID
* Company
* Country
* Job title
* Another common key

---

### 11. M Language

Power Query transformations are based on **M language**.

M was used to understand and create transformation logic behind Power Query operations.

Example:

```powerquery
= Table.SelectRows(Source, each [Salary] > 50000)
```

M language provides greater control over data transformation and allows repeatable transformation steps.

---

# 🧮 Power Pivot

Power Pivot was used for advanced data modeling and analysis.

It allows multiple tables to be connected and analyzed as part of a data model.

Power Pivot was used for:

* Data modeling
* Managing relationships
* Connecting multiple tables
* Creating calculated measures
* Handling larger datasets
* Building advanced analytical models

---

# 🧠 DAX

**DAX (Data Analysis Expressions)** was used with Power Pivot for advanced calculations and measures.

Examples of DAX functions include:

```text
SUM()
AVERAGE()
COUNT()
COUNTROWS()
DISTINCTCOUNT()
CALCULATE()
FILTER()
IF()
DIVIDE()
```

DAX can be used to create analytical measures such as:

* Total Jobs
* Average Salary
* Median/summary measures where appropriate
* Job Count
* Distinct Companies
* Salary comparisons
* Percentage calculations
* Filter-based calculations

Example:

```DAX
Total Jobs = COUNTROWS(JobData)
```

Another example:

```DAX
Average Salary = AVERAGE(JobData[Salary])
```

---

# 📊 Advanced Dashboard Development

The advanced features were combined to create the final analytical dashboard.

### Workflow

```text
Raw Data
   ↓
Power Query
   ↓
Data Cleaning
   ↓
Transformation
   ↓
Append / Merge
   ↓
Data Model
   ↓
Power Pivot
   ↓
DAX Measures
   ↓
Pivot Tables
   ↓
Pivot Charts
   ↓
Dashboard
   ↓
Business Insights
```

This workflow reduces repetitive manual work and creates a more structured approach to data analysis.

---

# 📌 Project 2 Skill Summary

```text
Advanced Microsoft Excel
│
├── Pivot Tables
├── Advanced Pivot Table Analysis
├── Pivot Charts
├── Analysis ToolPak
├── Statistical Analysis
│
├── Power Query
│   ├── Data Cleaning
│   ├── Data Transformation
│   ├── Append
│   ├── Merge
│   ├── Advanced Transformations
│   └── M Language
│
├── Power Pivot
│   ├── Data Modeling
│   ├── Relationships
│   └── Measures
│
└── DAX
    ├── Calculated Measures
    ├── Aggregations
    ├── Filtering
    └── Advanced Calculations
```

---

# ⚖️ Project 1 vs Project 2

| Area                    | Project 1 | Project 2 |
| ----------------------- | --------- | --------- |
| Microsoft Excel         | ✅         | ✅         |
| Logical Functions       | ✅         | ✅         |
| Mathematical Functions  | ✅         | ✅         |
| Statistical Functions   | ✅         | ✅         |
| Formatting              | ✅         | ✅         |
| Conditional Formatting  | ✅         | ✅         |
| Basic Charts            | ✅         | ✅         |
| Data Analysis           | ✅         | ✅         |
| Pivot Tables            | —         | ✅         |
| Advanced Pivot Tables   | —         | ✅         |
| Pivot Charts            | —         | ✅         |
| Analysis ToolPak        | —         | ✅         |
| Power Query             | —         | ✅         |
| Advanced Transformation | —         | ✅         |
| Append Queries          | —         | ✅         |
| Merge Queries           | —         | ✅         |
| M Language              | —         | ✅         |
| Power Pivot             | —         | ✅         |
| DAX                     | —         | ✅         |
| Data Modeling           | —         | ✅         |
| Advanced Dashboard      | —         | ✅         |

---

# 💼 Skills Demonstrated Across Both Projects

Together, these projects demonstrate a progression from **Excel fundamentals to advanced Excel-based data analytics and business intelligence**.

### Excel Fundamentals

* Excel Functions
* Logical Analysis
* Mathematical Calculations
* Statistical Analysis
* Formatting
* Conditional Formatting
* Charts
* Data Sorting & Filtering

### Advanced Excel

* Pivot Tables
* Pivot Charts
* Analysis ToolPak
* Power Query
* Data Transformation
* Append Queries
* Merge Queries
* M Language
* Power Pivot
* Data Modeling
* DAX

### Data Analytics

* Data Cleaning
* Data Transformation
* Exploratory Data Analysis
* Statistical Analysis
* Data Visualization
* Dashboard Development
* Business Insights
* Data Modeling

