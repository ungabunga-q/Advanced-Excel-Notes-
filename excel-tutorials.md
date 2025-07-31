# Excel Tutorials

## 1. Getting Started with Excel 🚀
- **How to open Excel:**  
  Launch from Start Menu or Run dialog (`Windows + R → excel`).  
- **Overview of the interface:**  
  Explore Ribbon, Tabs, Formula Bar, and Worksheet area.
- **Creating and saving a workbook:**  
  Click `File → New` to create, `File → Save As` to save.

## 2. Basic Data Entry ✍️
- **Entering text, numbers, and dates:**  
  Click a cell and type; press Enter to confirm.
- **Editing and deleting data:**  
  Double-click to edit, press Delete to clear.
- **Autofill and series:**  
  Drag fill handle to copy or create sequences.

## 3. Formatting Data 🎨
- **Font, color, and cell styles:**  
  Use Home tab to change font, add colors, and apply styles.
- **Number formats:**  
  Format cells as currency, date, percentage, etc.
- **Conditional formatting:**  
  Highlight cells based on rules (e.g., values above 100).

## 4. Formulas and Functions ➕
- **Writing basic formulas (`=A1+B1`):**  
  Start with `=`, use cell references and operators.
- **Using built-in functions (SUM, AVERAGE, COUNT):**  
  Type `=SUM(A1:A10)` to add values.
- **Relative vs. absolute references:**  
  `$A$1` is absolute, `A1` is relative.

## 5. Working with Tables 📋
- **Creating tables:**  
  Select data, click `Insert → Table`.
- **Sorting and filtering data:**  
  Use table headers to sort or filter.
- **Table design options:**  
  Change styles, add total rows, banded rows.

## 6. Charts and Visualizations 📊
- **Creating charts (Column, Pie, Line):**  
  Select data, click `Insert → Chart`.
- **Customizing chart elements:**  
  Edit titles, legends, colors, and data labels.

## 7. Data Analysis Tools 🔍
- **Pivot Tables:**  
  Summarize large data sets easily.
- **Data validation:**  
  Restrict input (e.g., only numbers or dates).
- **Removing duplicates:**  
  Clean up lists by deleting repeated entries.

## 8. Automating Tasks 🤖
- **Recording macros:**  
  Automate repetitive actions by recording steps.
- **Introduction to VBA:**  
  Write custom scripts for advanced automation.

## 9. Collaboration and Review 🤝
- **Adding comments:**  
  Right-click a cell, select `New Comment` to add notes.
- **Protecting sheets:**  
  Lock cells or sheets to prevent unwanted changes.
- **Tracking changes:**  
  Enable `Track Changes` to monitor edits.

---

## 10. Useful Excel Shortcuts ⌨️
- **Ctrl + C/V/X:** Copy/Paste/Cut
- **Ctrl + Z/Y:** Undo/Redo
- **Ctrl + S:** Save
- **Ctrl + F/H:** Find/Replace
- **Ctrl + Arrow Keys:** Jump to data edges

## 11. Common Excel Errors ⚠️
- **#DIV/0!:** Division by zero
- **#VALUE!:** Wrong data type
- **#REF!:** Invalid cell reference
- **#NAME?:** Unrecognized formula name


## 12. Excel Functions Reference 📚

Below is a list of commonly used Excel functions with explanations:

### Mathematical & Trigonometric Functions ➗
- **SUM(range):** Adds all numbers in a range.  
  *Example:* `=SUM(A1:A10)`
- **AVERAGE(range):** Calculates the mean of numbers.  
  *Example:* `=AVERAGE(B1:B5)`
- **COUNT(range):** Counts numeric cells.  
  *Example:* `=COUNT(C1:C10)`
- **COUNTA(range):** Counts non-empty cells.  
  *Example:* `=COUNTA(D1:D10)`
- **MAX(range):** Returns the largest value.  
  *Example:* `=MAX(E1:E10)`
- **MIN(range):** Returns the smallest value.  
  *Example:* `=MIN(F1:F10)`
- **ROUND(number, num_digits):** Rounds a number.  
  *Example:* `=ROUND(3.14159, 2)` → 3.14
- **SUMIF(range, criteria, [sum_range]):** Adds cells meeting criteria.  
  *Example:* `=SUMIF(A1:A10, ">100")`
- **SUMIFS(sum_range, criteria_range1, criteria1, ...):** Adds cells meeting multiple criteria.  
  *Example:* `=SUMIFS(B1:B10, A1:A10, ">100", C1:C10, "<50")`
- **PRODUCT(range):** Multiplies all numbers in a range.  
  *Example:* `=PRODUCT(G1:G5)`
- **MOD(number, divisor):** Returns remainder.  
  *Example:* `=MOD(10, 3)` → 1

### Logical Functions 🧠
- **IF(logical_test, value_if_true, value_if_false):** Returns one value if condition is true, another if false.  
  *Example:* `=IF(A1>10, "Yes", "No")`
- **AND(condition1, condition2, ...):** Returns TRUE if all conditions are true.  
  *Example:* `=AND(A1>0, B1<5)`
- **OR(condition1, condition2, ...):** Returns TRUE if any condition is true.  
  *Example:* `=OR(A1=1, B1=2)`
- **NOT(condition):** Reverses logical value.  
  *Example:* `=NOT(A1=5)`

### Text Functions 📝
- **CONCATENATE(text1, text2, ...):** Joins text strings.  
  *Example:* `=CONCATENATE("Hello ", "World")`
- **TEXT(value, format_text):** Formats a number as text.  
  *Example:* `=TEXT(1234.5, "$#,##0.00")`
- **LEFT(text, num_chars):** Returns leftmost characters.  
  *Example:* `=LEFT("Excel", 2)` → "Ex"
- **RIGHT(text, num_chars):** Returns rightmost characters.  
  *Example:* `=RIGHT("Excel", 3)` → "cel"
- **MID(text, start_num, num_chars):** Returns characters from middle.  
  *Example:* `=MID("Excel", 2, 3)` → "xce"
- **LEN(text):** Returns length of text.  
  *Example:* `=LEN("Excel")` → 5
- **UPPER(text):** Converts to uppercase.  
  *Example:* `=UPPER("excel")` → "EXCEL"
- **LOWER(text):** Converts to lowercase.  
  *Example:* `=LOWER("EXCEL")` → "excel"
- **TRIM(text):** Removes extra spaces.  
  *Example:* `=TRIM(" Excel ")` → "Excel"
- **REPLACE(old_text, start_num, num_chars, new_text):** Replaces part of text.  
  *Example:* `=REPLACE("Excel", 2, 3, "123")` → "E123l"
- **SUBSTITUTE(text, old_text, new_text):** Substitutes text.  
  *Example:* `=SUBSTITUTE("Excel", "cel", "lent")` → "Exellent"

### Date & Time Functions 🕒
- **TODAY():** Returns current date.
- **NOW():** Returns current date and time.
- **DATE(year, month, day):** Creates a date.
- **DAY(date):** Returns day of month.
- **MONTH(date):** Returns month.
- **YEAR(date):** Returns year.
- **HOUR(time):** Returns hour.
- **MINUTE(time):** Returns minute.
- **SECOND(time):** Returns second.
- **EDATE(start_date, months):** Adds months to date.
- **EOMONTH(start_date, months):** Returns last day of month after months.

### Lookup & Reference Functions 🔎
- **VLOOKUP(lookup_value, table_array, col_index, [range_lookup]):** Looks up value vertically.  
  *Example:* `=VLOOKUP("Apple", A2:B10, 2, FALSE)`
- **HLOOKUP(lookup_value, table_array, row_index, [range_lookup]):** Looks up value horizontally.
- **INDEX(array, row_num, [col_num]):** Returns value at position.
- **MATCH(lookup_value, lookup_array, [match_type]):** Finds position of value.
- **OFFSET(reference, rows, cols, [height], [width]):** Returns cell at offset.
- **INDIRECT(ref_text):** Returns reference from text.

### Statistical Functions 📊
- **MEDIAN(range):** Returns middle value.
- **MODE(range):** Returns most frequent value.
- **STDEV(range):** Estimates standard deviation.
- **VAR(range):** Estimates variance.
- **LARGE(range, k):** Returns k-th largest value.
- **SMALL(range, k):** Returns k-th smallest value.

### Information Functions ℹ️
- **ISBLANK(value):** Checks if cell is empty.
- **ISNUMBER(value):** Checks if value is number.
- **ISTEXT(value):** Checks if value is text.
- **ISERROR(value):** Checks for error.
- **TYPE(value):** Returns type of value.

---

> *Use these functions to analyze, manipulate, and automate your Excel data

---

> *For detailed steps, expand each section or request a