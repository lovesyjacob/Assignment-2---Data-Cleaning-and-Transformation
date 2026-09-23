# Assignment 2 – Data Cleaning and Transformation
  This is the second assignment in excel which handles  Data Cleaning and Transformation

## Objective
Perform data cleaning and transformation operations on the dataset in the **Answer** sheet using Excel.

# 1. Handling Missing Values
   
   ## 1.1 Missing Values in the Price Column

**Question:** Check for missing values in the `Price` column. How would you handle products with missing price information?

### Solution Applied

Three products had missing prices. The missing prices were handled using the **median price of the respective product category**.

| Product | Category | Cell | Applied Price |
|---|---|---|---:|
| Sony – Headphones | Electronics | F7 | $600 |
| Coleman – Camping Tent | Outdoor | F21 | $130 |
| Ray-Ban – Sunglasses | Fashion | F26 | $70 |

### Excel formulas used

**F7 – Sony Headphones:**
```excel
=MEDIAN(F2,F5,F10,F12,F15,F16,F18,F19,F22,F24,F25,F30)
```
Result: **600**

**F21 – Coleman Camping Tent:**
```excel
=MEDIAN(F11,F21)
```
Result: **130**

**F26 – Ray-Ban Sunglasses:**
```excel
=MEDIAN(F3,F8,F13,F20,F28,F31)
```
Result: **70**

The calculated results were then entered into the corresponding missing Price cells.

**Reason:** Median was used because it provides a representative value.

---

## 1.2 Missing Values in the Category Column

**Question:** If there are products with missing categories, propose a strategy to impute or deal with these missing values effectively.

### Solution Applied

Four missing category values were identified and filled based on similar products/product types already present in the dataset.

| Product | Brand | Cell | Category Applied |
|---|---|---|---|
| Backpack | North Face | F6 | Accessories |
| Sneakers | Adidas | F13 | Fashion |
| Coffee Maker | Nespresso | F14 | Kitchen |
| Fitness Tracker | Xiaomi | F24 | Electronics |

**Strategy:** Missing categories were assigned based on similar products and the existing category structure in the dataset. If a category cannot be reliably determined, it should be verified from the source rather than guessed.

---
# 2. Correcting Inconsistent Data

## 2.1 Inconsistent Product Name Formats

**Question:** Identify any inconsistent text formats present in the `Product Name` column.

### Solution Applied

The following inconsistent lowercase product names were identified:

| Original | Standardized |
|---|---|
| laptop | Laptop |
| smartphone | Smartphone |
| headphones | Headphones |

The **Find and Replace** feature (`Ctrl + H`) was used to standardize the values. Also can use find&select option from the formula bar

---

## 2.2 Typos in the Category Column

**Question:** Identify any typos present in the `Category` column.

### Solution Applied

The typo:

```text
Electroni
```

was corrected to:

```text
Electronics
```

using **Find and Replace**.

---

## 2.3 Standardizing Product Names and Categories

**Question:** Use the Find and Replace function to standardize the text formats in the `Product Name` column and fix typos or misspellings in the `Category` column.

### Solution Applied

The following replacements were performed:

- `laptop` → `Laptop`
- `smartphone` → `Smartphone`
- `headphones` → `Headphones`
- `Electroni` → `Electronics`
---

# 3. Removing Duplicates

**Question:** Identify any duplicate rows within the dataset based on the entirety of each row, and remove them if any.

### Solution Applied

The complete dataset was selected and **Data → Remove Duplicates** was used with all columns selected.

Three complete duplicate rows were removed:

- Laptop Bag – Samsonite
- Headphones – Bose
- Laptop – HP

The dataset changed from **34 original data rows to 31 unique data rows**.

---

# 4. Splitting and Merging Data

## 4.1 Split Product ID

**Question:** Split the `Product ID` column into two separate columns for `Manufacturing Date` and `Country Code`. Remove unnecessary characters, if any.

### Solution Applied

The original Product ID follows the pattern:

```text
28-JAN-US
```

It contains the day/month and country code. Two new columns were created:

- `Manufacturing Date`
- `Country Code`

Because the original Product ID did not contain a year, **2026 was used as the default year**.

### Manufacturing Date formula

For Product ID in A2:

```excel
=DATE(2026,MATCH(MID(A2,4,3),{"JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"},0),VALUE(LEFT(A2,2)))
```

For example:

```text
28-JAN-US
```

becomes a valid Excel date representing:

```text
28-Jan-2026
```

and is displayed as:

```text
28-01-2026
```

### Formula explanation

- `LEFT(A2,2)` extracts the day.
- `VALUE()` converts the day from text to a number.
- `MID(A2,4,3)` extracts the month abbreviation.
- `MATCH()` converts the month abbreviation into its month number.
- `DATE()` creates a valid Excel date using 2026 as the year.

### Country Code formula

```excel
=RIGHT(A2,2)
```

For `28-JAN-US`, the result is `US`.

---

## 4.2 Merge Brand Name and Product Name

**Question:** Merge the `Brand Name` and `Product Name` columns into one column named `Product Brand`.

### Solution Applied

A new `Product Brand` column was created using:

```excel
=E2&" "&D2
```

For example:

```text
Dell + Laptop
```

becomes:

```text
Dell Laptop
```

The `" "` adds a space between the two values.

---

# 5. Number Formatting

## 5.1 Price Currency Format

**Question:** Format the data type of the `Price` column to currency format.

### Solution Applied

The `Price ($)` column was formatted using Excel's **Currency** format.

Example:

```text
1000
```

is displayed as:

```text
$1,000.00
```

No formula was required because this is a formatting operation.

---

## 5.2 Manufacturing Date Format

**Question:** Format the `Manufacturing Date` column to display dates in `DD-MM-YYYY` format.

### Solution Applied

The Manufacturing Date was first converted into a valid Excel date using the `DATE()` formula described above.

The column was then formatted as:

```text
dd-mm-yyyy
```

Example:

```text
28-Jan-2026
```

is displayed as:

```text
28-01-2026
```

---

# 6. Conditional Formatting

## 6.1 Price Column

**Question:** Apply data bar or color scale conditional formatting in the `Price` column.

### Solution Applied

**Home → Conditional Formatting → Data Bars → Gradient Fill** was applied to the `Price ($)` column.

Higher prices are represented by longer data bars.

No formula was required.

---

## 6.2 Category Column

**Question:** Create a custom rule for conditional formatting in the `Category` column to highlight cells where the category is `Electronics`.

### Solution Applied

A conditional formatting rule was created for the `Category` column:

**Home → Conditional Formatting → New Rule → Format only cells that contain**

Condition:

```text
Cell Value = "Electronics"
```
A fill format was selected to highlight matching cells.
No formula was required for this rule.

---

# Final Summary

The following operations were completed:

- Missing Price values identified and filled using category median values.
- Missing Category values identified and filled using similar product/category information.
- Inconsistent Product Name capitalization corrected using Find and Replace.
- Category typo `Electroni` corrected to `Electronics`.
- Duplicate rows removed using Excel's Remove Duplicates feature.
- Product ID split into Manufacturing Date and Country Code.
- Default year **2026** added because the source Product ID did not contain a year.
- Manufacturing Date converted to a valid Excel date.
- Manufacturing Date displayed in `DD-MM-YYYY` format.
- Brand Name and Product Name merged into Product Brand.
- Price formatted as currency.
- Data Bars applied to the Price column.
- Electronics category highlighted using conditional formatting.

## Additional Notes

The assignment  required Excel's `MEDIAN()` fn used for the missing-price calculations, while `DATE()`, `MATCH()`, `MID()`, `VALUE()`, `LEFT()`, and `RIGHT()` were used for the Product ID transformation.

**Original dataset:** 34 data rows  
**After duplicate removal:** 31 unique data rows
