# Advanced Excel Notes

## Table of Contents
1. [Introduction to Microsoft Excel](#introduction-to-microsoft-excel)
2. [How to Start MS Excel](#how-to-start-ms-excel)
3. [Microsoft Excel Interface Overview](#microsoft-excel-interface-overview)
4. [Home Tab Features](#home-tab-features)
5. [Insert Tab Features](#insert-tab-features)
6. [Page Layout Tab](#page-layout-tab)
7. [Formulas Tab](#formulas-tab)
8. [Data Tab](#data-tab)
9. [Review Tab](#review-tab)
10. [View Tab](#view-tab)
11. [Macros](#macros)

---

## Introduction to Microsoft Excel

Microsoft Excel is a powerful spreadsheet software developed by Microsoft. It enables users to organize, format, and calculate data using formulas within a tabular format (rows & columns).

---

## How to Start MS Excel

- **Start Menu Method:**  
  `Start → All Programs → Microsoft Office → Microsoft Excel`
- **Run Dialog Method:**  
  `Windows + R → Type "excel" → OK`

**Excel Tabs:**  
- Home
- Insert
- Page Layout
- Formulas
- Data
- Review
- View

---

## Microsoft Excel Interface Overview

| Element                 | Description                                                                 |
|-------------------------|-----------------------------------------------------------------------------|
| Quick Access Toolbar    | Shortcuts for frequently used tasks                                         |
| Tabs                    | Shortcuts to accomplish specific tasks                                      |
| Cell                    | Intersection of a row and column                                            |
| Ribbon                  | Main location for commands                                                  |
| Name Box                | Displays active cell address                                                |
| Formula Bar             | Shows cell content or formula                                               |
| Row                     | Numbers (1 - 1048576) on worksheet's left side                             |
| Column                  | Letters (A - XFD) on worksheet's top                                       |
| Cell Address            | Combination of column letter and row number                                 |
| Worksheet Views         | Different views: Normal, Page Layout, Page Break View, etc.                 |
| Zoom                    | Magnifies worksheet for better viewing                                      |
| Worksheet Tab           | Represents different sheets in a workbook                                   |

---

## Home Tab Features

### Clipboard
- **Paste:** Insert copied/cut data  
- **Cut:** Remove selection and place on clipboard  
- **Copy:** Duplicate selection to clipboard  
- **Format Painter:** Copy formatting to other areas

### Font
- **Font, Font Size:** Change font style and size
- **Bold, Italic, Underline:** Emphasize text
- **Borders:** Add cell borders
- **Fill Color:** Background color for cells
- **Font Color:** Change font color

### Alignment
- **Top, Middle, Bottom:** Vertical alignment
- **Left, Center, Right:** Horizontal alignment
- **Orientation:** Rotate text
- **Indent:** Increase/Decrease space from cell border

### Merge and Center
- **Merge & Center:** Combine cells and center text
- **Merge Across:** Merge cells horizontally
- **Merge Cells:** Merge selected cells
- **Unmerge Cells:** Undo cell merging

### Wrap Text
- **Wrap Text:** Display text on multiple lines within a cell

### Number Formatting
- **Formats:** General, Number, Currency, Accounting, Date, Time, Percentage, Fraction, Scientific

### Styles
- **Conditional Formatting:** Highlight data by rules
- **Format as Table:** Quick table styles
- **Cell Styles:** Apply consistent formats

### Cell References
- **Absolute:** $A$1 (row & column fixed)
- **Relative:** A1 (both change when copied)
- **Mixed:** $A1 or A$1 (either row or column fixed)

### Cells
- **Insert:** Add cells, rows, columns, sheets
- **Delete:** Remove cells, rows, columns, sheets
- **Format:** Change dimensions, protect, hide

### Editing
- **AutoSum:** Automatically add values
- **Fill:** Continue patterns into cells
- **Clear:** Remove contents, formatting, comments, hyperlinks
- **Find & Replace:** Search and replace text/content
- **Go To:** Jump to specific location in workbook

### Sort and Filter
- **Sort:**  
  - A to Z: Ascending  
  - Z to A: Descending  
  - Custom Sort
- **Filter:**  
  - View specific rows, hide others  
  - Types: Text, Numbers, Dates, Color

---

## Insert Tab Features

### Pivot Table
- Summarize and organize data by rows and columns
- **Slicer:** Interactive filtering control for Pivot Tables
- **Timeline:** Select period of dates for Pivot Table

### Tables
- Organize data and access design options
- **Design:** Name, resize, summarize, convert to range
- **Table Style Options:** Header row, total row, banded rows/columns, first/last column
- **Quick Styles:** Change visual style

### Illustrations
- **Pictures:** Insert images from device, with formatting options (remove background, corrections, artistic effects, compress, change/reset picture, borders/effects/layout, arrange, crop, size)
- **Online Pictures:** Download images from online sources
- **Shapes:** Add ready-made shapes, format with fill, outline, effects, arrange, crop, size
- **SmartArt:** Visual graphics for information
  - Design: Add shape/bullet, text pane, promote/demote, move, layout, style, color, reset
  - Format: Change shape, size, fill, outline, effects, text style, arrange, crop, size
- **Screenshot:** Add snapshots of open windows

### Charts
- Visual representation of data
- **Chart Types:** Column, Line, Pie, Doughnut, Bar, Area, Scatter, Bubble, Stock, Surface, Radar, Combo, Hierarchy, Statistical, Waterfall, Histogram, Box & Whisker, Tree Map, Sunburst

### Sparklines
- Tiny charts inside cells showing trends
- **Types:** Line, Column, Win-Loss

### Links
- **Hyperlink:** Reference to file, webpage, document location, new document, email
- **Bookmark:** Jump to specific document place
- **Cross Reference:** Refer to headings, figures, tables

### Text
- **Text Box:** Frame for entering text
- **Header & Footer:** Repeat content at page top/bottom
- **WordArt:** Artistic text styles
- **Signature:** Insert signature line
- **Object:** Excel entities (Worksheets, Rows, Columns, Ranges, Workbook)

### Symbols
- **Equation:** Add mathematical equation
- **Symbol:** Add special characters

---

## Page Layout Tab

### Themes
- Changes overall formatting (font, size, colors)

### Page Setup
- **Margin:** Set printable area
- **Orientation:** Landscape or Portrait
- **Page Size:** Paper size
- **Print Area:** Select cells to print
- **Page Break:** Define where printed page ends
- **Background:** Sheet background (not printed)
- **Print Titles:** Repeats row/column headings on every page

### Scale to Fit
- **Width/Height:** Fit printout to pages
- **Scale:** Stretch/Shrink to percentage

### Sheet Options
- **Gridlines:** Show/print gridlines
- **Headings:** Show/print row/column headers

### Arrange
- **Bring Forward/Send Backward:** Change object layering
- **Selection Pane:** List of objects
- **Align/Group/Rotate:** Change object placement/format

---

## Formulas Tab

### Formulas
- User-written statements for calculations; all start with `=`

### Financial Functions
- FV, FVSCHEDULE, PV, NPV, XNPV, PMT, PPMT, IRR, MIRR, XIRR, NPER, RATE, EFFECT, NOMINAL, SLN

### Logical Functions
- AND, OR, XOR, NOT, FALSE, IF, IFERROR, IFNA

### Text Functions
- CONCATENATE, DOLLAR, EXACT, FIND, FIXED, LEFT, LEN, LOWER, MID, PROPER, REPLACE, REPT, RIGHT, SEARCH, SUBSTITUTE, TEXT, TEXTJOIN, TRIM, UPPER

### Date & Time Functions
- DATE, DAY, DAYS, DAYS360, EDATE, EOMONTH, HOUR, MINUTE, MONTH, NOW, SECOND, TIME, TODAY, WEEKDAY, WEEKNUM, WORKDAY, YEAR

### Lookup Functions
- ADDRESS, AREAS, CHOOSE, COLUMN, COLUMNS, FORMULATEXT, GETPIVOTTABLE, HLOOKUP, HYPERLINK, INDEX, INDIRECT, LOOKUP, MATCH, OFFSET, ROW, ROWS, TRANSPOSE, VLOOKUP

### Math & Trig Functions
- ABS, AGGREGATE, EVEN, FACT, GCD, INT, LCM, MOD, ODD, PI, POWER, PRODUCT, QUOTIENT, RAND, RANDBETWEEN, ROMAN, ROUND, ROUNDUP, ROUNDDOWN, SQRT, SQRTPI, SUM, SUMIF, SUMIFS, SUMPRODUCT, SUMSQ

### Statistical Functions
- AVERAGE, MODE, AVERAGEA, AVERAGEIF, AVERAGEIFS, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, LARGE, MAX, MAXA, MAXIFS, MEDIAN, MIN, MINIFS, MINA, SMALL

### Information Functions
- CELL, ISBLANK, ISERR, ISERROR, ISEVEN, ISFORMULA, ISLOGICAL, ISNA, ISNONTEXT, ISNUMBER, ISODD, ISREF, ISTEXT, N, NA, SHEET, SHEETS, TYPE

### Defined Names
- **Name Manager:** Manage defined names/tables
- **Defined Name:** Create and apply names
- **Use in Formula:** Insert names into formulas
- **Create from Selection:** Auto-generate names from selected cells

### Formula Auditing
- **Trace Precedents/Dependents:** Show cell relationships
- **Remove Arrows:** Clear auditing arrows
- **Show Formula:** Display formulas instead of values
- **Error Checking:** Find formula errors
- **Evaluate Formula:** Debug formulas in steps
- **Watch Window:** Monitor cell values

### Calculation
- **Options:** Automatic/Manual calculation
- **Calculate Now/Sheet:** Trigger recalculation

---

## Data Tab

### Get External Data
- Import from Access, Web, Text, Other sources

### Get & Transform
- **New Query:** Combine multiple data sources
- **Show Queries:** View queries list
- **From Table:** Create query from table
- **Recent Sources:** Manage recent sources

### Connections
- **Refresh All:** Update all data sources
- **Connections:** List all data connections

### Sort & Filter
- **Sort A-Z/Z-A:** Ascending/descending sort
- **Filter:** Display only matching data
- **Advanced Filter:** Complex criteria filtering

### Data Tools
- **Text to Columns:** Split text into columns
- **Flash Fill:** Auto-fill values
- **Remove Duplicates:** Delete duplicate rows
- **Data Validation:** Control cell input
- **Consolidate:** Summarize data from ranges
- **Relationships:** Link tables for reports
- **Manage Data Model:** Prepare and maintain data

### Forecast
- **Scenario Manager:** Create/switch between value groups
- **Goal Seek:** Find required input for desired output
- **Data Table:** See results for multiple inputs
- **Forecast Sheet:** Predict data trends

### Outline
- **Group/Ungroup:** Manage cell groups
- **Subtotal:** Calculate totals for related data

### Analyse
- **Solver:** Optimize and simulate models

---

## Review Tab

### Proofing
- **Spelling & Grammar:** Check document
- **Thesaurus:** Find synonyms and concepts

### Insights
- **Smart Lookup:** Extra information from the web

### Comments
- **New, Delete, Previous/Next, Show/Hide, Show All, Show Link:** Manage cell comments

### Changes
- **Protect Sheet/Workbook:** Restrict changes
- **Share Workbook:** Collaboration
- **Protect and Share Workbook:** Track changes with password
- **Allow Edit Ranges:** Password protection for ranges
- **Track Changes:** Monitor modifications

---

## View Tab

### Workbook Views
- **Normal, Page Break Preview, Page Layout, Custom Views:** Display options

### Show
- **Ruler, Formula Bar, Gridlines, Headings:** Toggle displays

### Zoom
- **Zoom, 100%, Zoom to Selection:** Magnification options

### Window
- **New Window, Arrange All, Freeze Panes, Split, Hide/Unhide, Side by Side, Synchronous Scrolling, Reset Position, Switch Windows:** Manage document windows

---

## Macros

- **Definition:** Macros automate repetitive tasks in Excel; written in VBA (Visual Basic for Applications).
- **File Extension:** `.MAC`
- **Types:**
  - **Record Macros:** Automate actions by recording.
  - **Code Macros (VBA):** Write custom automation scripts.

---

