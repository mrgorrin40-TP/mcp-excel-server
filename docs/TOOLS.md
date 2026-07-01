# Tools Reference

Complete reference for all MCP Excel Server tools.

## Table of Contents

- [Reading Tools](#reading-tools)
- [Writing Tools](#writing-tools)
- [Formula Tools](#formula-tools)
- [Analysis Tools](#analysis-tools)
- [Chart Tools](#chart-tools)
- [Table Tools](#table-tools)

---

## Reading Tools

### read_cell

Read a single cell value from an Excel file.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `cell` | string | Yes | Cell reference (e.g., "A1", "B2") |

**Returns:**
- Cell value (string, number, boolean, or null)
- Cell data type information

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "cell": "A1"
}
```

**Response:**
```json
{
  "value": "Product",
  "type": "string"
}
```

---

### read_range

Read multiple cells from an Excel file.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `range` | string | Yes | Range notation (e.g., "A1:C10") |
| `page_size` | integer | No | Rows per page (default: 100) |
| `page` | integer | No | Page number (default: 1) |

**Returns:**
- 2D array of cell values
- Pagination info if data exceeds page_size

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "range": "A1:C100",
  "page_size": 50
}
```

**Response:**
```json
{
  "data": [
    ["Product", "Quantity", "Price"],
    ["Widget A", 100, 29.99],
    ["Widget B", 50, 49.99]
  ],
  "total_rows": 100,
  "page": 1,
  "total_pages": 2
}
```

---

### get_sheet_info

Get metadata about a worksheet.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |

**Returns:**
- Column names and data types
- Row and column counts
- Used range
- Sample data

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1"
}
```

**Response:**
```json
{
  "name": "Q1",
  "columns": [
    {"name": "Product", "type": "string", "letter": "A"},
    {"name": "Quantity", "type": "number", "letter": "B"},
    {"name": "Price", "type": "number", "letter": "C"}
  ],
  "row_count": 100,
  "column_count": 3,
  "used_range": "A1:C101",
  "sample_data": [
    ["Widget A", 100, 29.99],
    ["Widget B", 50, 49.99]
  ]
}
```

---

### search_cells

Search for values in an Excel file.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `query` | string | Yes | Search term |
| `sheet_name` | string | No | Specific sheet (searches all if omitted) |
| `max_results` | integer | No | Maximum results (default: 50) |

**Returns:**
- List of matching cells
- Cell locations (sheet, cell, row, column)
- Matched values

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "query": "Widget A",
  "sheet_name": "Q1"
}
```

**Response:**
```json
{
  "matches": [
    {
      "sheet": "Q1",
      "cell": "A2",
      "row": 2,
      "column": "A",
      "value": "Widget A"
    },
    {
      "sheet": "Q1",
      "cell": "A15",
      "row": 15,
      "column": "A",
      "value": "Widget A"
    }
  ],
  "total_matches": 2
}
```

---

### list_sheets

List all worksheets in an Excel file.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |

**Returns:**
- List of worksheet names
- Active sheet indicator
- Sheet count

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx"
}
```

**Response:**
```json
{
  "sheets": ["Q1", "Q2", "Summary", "Charts"],
  "active_sheet": "Q1",
  "count": 4
}
```

---

### describe_workbook

Get comprehensive overview of an Excel workbook.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |

**Returns:**
- File metadata (size, modified date)
- All sheet names and summaries
- Total cell counts
- Data structure overview

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx"
}
```

**Response:**
```json
{
  "file": {
    "name": "sales.xlsx",
    "path": "C:/Documents/sales.xlsx",
    "size_kb": 125,
    "modified": "2024-01-15T10:30:00"
  },
  "sheets": [
    {
      "name": "Q1",
      "rows": 100,
      "columns": 5,
      "used_range": "A1:E101"
    },
    {
      "name": "Q2",
      "rows": 100,
      "columns": 5,
      "used_range": "A1:E101"
    }
  ],
  "total_sheets": 2,
  "total_cells": 1010
}
```

---

## Writing Tools

### write_cells

Write values to cells in an Excel file.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `range` | string | Yes | Target range (e.g., "A1:C3") |
| `values` | array[array] | Yes | 2D array of values to write |

**Returns:**
- Success status
- Cells written count

**Example:**
```json
{
  "file_path": "C:/Documents/report.xlsx",
  "sheet_name": "Summary",
  "range": "A1:C3",
  "values": [
    ["Name", "Score", "Grade"],
    ["Alice", 95, "A"],
    ["Bob", 85, "B"]
  ]
}
```

**Response:**
```json
{
  "success": true,
  "cells_written": 6,
  "range": "A1:C3"
}
```

---

### write_formula

Add an Excel formula to a cell.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `cell` | string | Yes | Target cell (e.g., "D1") |
| `formula` | string | Yes | Excel formula (with = prefix) |

**Returns:**
- Success status
- Formula that was written

**Example:**
```json
{
  "file_path": "C:/Documents/report.xlsx",
  "sheet_name": "Summary",
  "cell": "D1",
  "formula": "=SUM(B2:B10)"
}
```

**Response:**
```json
{
  "success": true,
  "cell": "D1",
  "formula": "=SUM(B2:B10)"
}
```

---

### create_sheet

Create a new worksheet in an Excel file.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Name for new worksheet |

**Returns:**
- Success status
- Sheet name created

**Example:**
```json
{
  "file_path": "C:/Documents/report.xlsx",
  "sheet_name": "Analysis"
}
```

**Response:**
```json
{
  "success": true,
  "sheet_name": "Analysis"
}
```

---

### delete_sheet

Delete a worksheet from an Excel file.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name to delete |

**Returns:**
- Success status
- Deleted sheet name

**Example:**
```json
{
  "file_path": "C:/Documents/report.xlsx",
  "sheet_name": "TempSheet"
}
```

**Response:**
```json
{
  "success": true,
  "sheet_name": "TempSheet",
  "message": "Sheet 'TempSheet' deleted successfully"
}
```

**Notes:**
- Cannot delete the last remaining worksheet
- Sheet must exist in the workbook

---

## Formula Tools

### read_formula

Read the formula text from a cell (returns formula text, not calculated value).

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `cell` | string | Yes | Cell reference (e.g., "D1") |

**Returns:**
- Cell location
- Formula text (if cell contains formula)
- Whether cell contains a formula

**Example:**
```json
{
  "file_path": "C:/Documents/report.xlsx",
  "sheet_name": "Summary",
  "cell": "D1"
}
```

**Response (cell has formula):**
```json
{
  "success": true,
  "cell": "D1",
  "formula": "=SUM(B2:B10)",
  "is_formula": true,
  "sheet_name": "Summary"
}
```

**Response (cell has no formula):**
```json
{
  "success": true,
  "cell": "A1",
  "value": "Product",
  "is_formula": false,
  "sheet_name": "Summary"
}
```

---

### get_formula_templates

Get common Excel formula templates.

**Parameters:** None

**Returns:**
- Dictionary of formula templates
- Usage instructions

**Example:**
```json
{}
```

**Response:**
```json
{
  "success": true,
  "templates": {
    "sum": "=SUM({range})",
    "average": "=AVERAGE({range})",
    "count": "=COUNT({range})",
    "if": "=IF({condition},\"{true_value}\",\"{false_value}\")",
    "vlookup": "=VLOOKUP({lookup_value},{table_range},{col_index},FALSE)"
  },
  "usage": "Use these templates as a starting point. Replace {range}, {criteria}, etc. with actual values."
}
```

---

## Analysis Tools

### get_column_stats

Get statistical summary of a column.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `column` | string | Yes | Column letter or name |

**Returns:**
- Count, sum, mean, median
- Min, max, range
- Standard deviation, variance
- Null count
- Data type

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "column": "B"
}
```

**Response:**
```json
{
  "column": "B",
  "name": "Quantity",
  "count": 100,
  "sum": 5000,
  "mean": 50.0,
  "median": 45.0,
  "min": 5,
  "max": 200,
  "range": 195,
  "std_dev": 35.5,
  "variance": 1260.25,
  "nulls": 2,
  "type": "number"
}
```

---

### filter_rows

Filter rows based on conditions.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `filters` | array[object] | Yes | Filter conditions |
| `logic` | string | No | "AND" or "OR" (default: "AND") |

**Filter Object:**
```json
{
  "column": "B",
  "operator": ">",
  "value": 100
}
```

**Operators:**
- Comparison: `==`, `!=`, `>`, `<`, `>=`, `<=`
- Text: `contains`, `startswith`, `endswith`
- Set: `in`, `not_in`
- Null: `is_null`, `is_not_null`

**Returns:**
- Filtered rows
- Row count

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "filters": [
    {"column": "B", "operator": ">", "value": 100},
    {"column": "C", "operator": ">=", "value": 25.00}
  ],
  "logic": "AND"
}
```

**Response:**
```json
{
  "data": [
    ["Widget A", 150, 29.99],
    ["Widget C", 120, 39.99]
  ],
  "row_count": 2,
  "total_rows": 100
}
```

---

### group_by

Group data and apply aggregation.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `columns` | array[string] | Yes | Columns to group by |
| `agg_column` | string | Yes | Column to aggregate |
| `agg_func` | string | Yes | Aggregation function |

**Aggregation Functions:**
- `sum` - Total
- `mean` - Average
- `median` - Middle value
- `count` - Count of rows
- `min` - Minimum
- `max` - Maximum
- `std` - Standard deviation

**Returns:**
- Grouped data
- Aggregated values

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "columns": ["Region"],
  "agg_column": "Sales",
  "agg_func": "sum"
}
```

**Response:**
```json
{
  "groups": [
    {"Region": "North", "Sales_sum": 50000},
    {"Region": "South", "Sales_sum": 45000},
    {"Region": "East", "Sales_sum": 60000},
    {"Region": "West", "Sales_sum": 55000}
  ],
  "group_count": 4
}
```

---

## Chart Tools

### create_chart

Create a chart in an Excel worksheet.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `chart_type` | string | Yes | Chart type: bar, line, pie, scatter, area |
| `title` | string | Yes | Chart title |
| `data_range` | string | Yes | Data range (e.g., "A1:D10") |
| `category_range` | string | No | Category labels range (e.g., "A1:A10") |
| `position` | string | No | Position (e.g., "E2" or "E2:J20") |

**Returns:**
- Chart type
- Title
- Position

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "chart_type": "bar",
  "title": "Sales by Product",
  "data_range": "A1:B10",
  "category_range": "A1:A10"
}
```

**Response:**
```json
{
  "success": true,
  "chart_type": "bar",
  "title": "Sales by Product",
  "sheet_name": "Q1",
  "position": "E2",
  "data_range": "A1:B10"
}
```

---

### list_charts

List all charts in an Excel worksheet.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |

**Returns:**
- List of charts with index, title, type, and position

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1"
}
```

**Response:**
```json
{
  "success": true,
  "sheet_name": "Q1",
  "charts": [
    {
      "index": 0,
      "title": "Sales by Product",
      "type": "BarChart",
      "anchor": "E2"
    }
  ],
  "count": 1
}
```

---

### modify_chart

Modify chart properties in an Excel worksheet.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `chart_index` | integer | Yes | Index of chart to modify (0-based) |
| `title` | string | No | New chart title |
| `style` | integer | No | Chart style (1-48) |

**Returns:**
- Modified properties list

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "chart_index": 0,
  "title": "Updated Title"
}
```

**Response:**
```json
{
  "success": true,
  "sheet_name": "Q1",
  "chart_index": 0,
  "modified_properties": ["title"]
}
```

---

### delete_chart

Delete a chart from an Excel worksheet.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `chart_index` | integer | Yes | Index of chart to delete (0-based) |

**Returns:**
- Deleted chart info

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "chart_index": 0
}
```

**Response:**
```json
{
  "success": true,
  "sheet_name": "Q1",
  "deleted_chart_index": 0,
  "chart_title": "Sales by Product"
}
```

---

## Table Tools

### create_table

Create a structured Excel table from a data range.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `table_name` | string | Yes | Name for the table |
| `data_range` | string | Yes | Data range (e.g., "A1:D100") |
| `has_header` | boolean | No | Whether data has headers (default: true) |
| `style` | string | No | Table style (e.g., "TableStyleMedium2") |

**Returns:**
- Table name
- Data range
- Style applied

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "table_name": "SalesTable",
  "data_range": "A1:D100"
}
```

**Response:**
```json
{
  "success": true,
  "table_name": "SalesTable",
  "sheet_name": "Q1",
  "data_range": "A1:D100",
  "has_header": true,
  "style": "TableStyleMedium2"
}
```

---

### list_tables

List all structured tables in an Excel worksheet.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |

**Returns:**
- List of tables with name, display name, range, and style

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1"
}
```

**Response:**
```json
{
  "success": true,
  "sheet_name": "Q1",
  "tables": [
    {
      "name": "Table1",
      "display_name": "SalesTable",
      "ref": "A1:D100",
      "style": "TableStyleMedium2"
    }
  ],
  "count": 1
}
```

---

### delete_table

Delete a structured table from an Excel worksheet.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `table_name` | string | Yes | Name of table to delete |

**Returns:**
- Deleted table name

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "table_name": "SalesTable"
}
```

**Response:**
```json
{
  "success": true,
  "sheet_name": "Q1",
  "deleted_table": "SalesTable"
}
```

---

### add_table_row

Add a row to an Excel structured table.

**Parameters:**

| Name | Type | Required | Description |
|------|------|----------|-------------|
| `file_path` | string | Yes | Absolute path to Excel file |
| `sheet_name` | string | Yes | Worksheet name |
| `table_name` | string | Yes | Table name |
| `values` | array | Yes | Row values |

**Returns:**
- Added values

**Example:**
```json
{
  "file_path": "C:/Documents/sales.xlsx",
  "sheet_name": "Q1",
  "table_name": "SalesTable",
  "values": ["Widget X", 50, 29.99, "North"]
}
```

**Response:**
```json
{
  "success": true,
  "sheet_name": "Q1",
  "table_name": "SalesTable",
  "values": ["Widget X", 50, 29.99, "North"]
}
```

---

## Coming Soon

The following tools are planned for future releases:

### Advanced (v0.3.0)
- `pivot_table` - Create pivot tables
- `conditional_format` - Apply conditional formatting
- `data_validation` - Add data validation rules
- `protect_sheet` - Protect worksheet with password
