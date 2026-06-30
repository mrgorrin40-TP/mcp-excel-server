# AGENTS.md - Guide for AI Agents

This document provides guidance for AI agents interacting with the MCP Excel Server.

## Overview

MCP Excel Server exposes Excel file manipulation capabilities through the Model Context Protocol. Use the available tools to read, write, analyze, and transform Excel files (.xlsx).

## Key Principles

### 1. Always Use Absolute Paths

```
✓ CORRECT: "C:/Documents/data.xlsx"
✓ CORRECT: "/home/user/documents/data.xlsx"
✗ WRONG: "data.xlsx"
✗ WRONG: "../data.xlsx"
```

### 2. Check Before Writing

Always verify file and sheet existence before writing:

```
1. Call describe_workbook() to verify file exists
2. Call list_sheets() to see available worksheets
3. Then proceed with write operations
```

### 3. Use Pagination for Large Data

Never read entire large files at once:

```
✓ CORRECT: read_range(range="A1:J100", page_size=100)
✗ WRONG: read_range(range="A1:Z10000")
```

## Tool Reference

### Reading Tools

#### read_cell
Read a single cell value.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - sheet_name: str (required) - Worksheet name
  - cell: str (required) - Cell reference (e.g., "A1")

Returns: Cell value (string, number, or null)
```

#### read_range
Read multiple cells in a range.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - sheet_name: str (required) - Worksheet name
  - range: str (required) - Range notation (e.g., "A1:C10")
  - page_size: int (optional) - Rows per page (default: 100)

Returns: 2D array of values
```

#### get_sheet_info
Get metadata about a worksheet.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - sheet_name: str (required) - Worksheet name

Returns: Column names, types, row count, sample data
```

#### search_cells
Search for values in a worksheet.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - query: str (required) - Search term
  - sheet_name: str (optional) - Specific sheet or all sheets

Returns: List of matching cells with locations
```

### Writing Tools

#### write_cells
Write values to cells.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - sheet_name: str (required) - Worksheet name
  - range: str (required) - Target range
  - values: list[list] (required) - 2D array of values

Returns: Success status
```

#### write_formula
Add an Excel formula.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - sheet_name: str (required) - Worksheet name
  - cell: str (required) - Target cell
  - formula: str (required) - Excel formula (with = prefix)

Returns: Success status
```

#### create_sheet
Create a new worksheet.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - sheet_name: str (required) - Name for new sheet

Returns: Success status
```

### Analysis Tools

#### get_column_stats
Get statistical summary of a column.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - sheet_name: str (required) - Worksheet name
  - column: str (required) - Column letter or name

Returns: Count, mean, median, min, max, std, nulls
```

#### filter_rows
Filter rows by conditions.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - sheet_name: str (required) - Worksheet name
  - filters: list[dict] (required) - Filter conditions
  - logic: str (optional) - "AND" or "OR" (default: "AND")

Returns: Filtered rows
```

#### group_by
Group and aggregate data.

```
Parameters:
  - file_path: str (required) - Absolute path to Excel file
  - sheet_name: str (required) - Worksheet name
  - columns: list[str] (required) - Grouping columns
  - agg_column: str (required) - Column to aggregate
  - agg_func: str (required) - Function: sum, mean, median, count, min, max

Returns: Aggregated results
```

## Common Workflows

### Workflow 1: Explore Unknown File

```
1. describe_workbook(file_path) - Get overview
2. list_sheets(file_path) - See all worksheets
3. get_sheet_info(file_path, sheet_name) - Understand structure
4. read_range(file_path, sheet_name, "A1:Z10") - Sample data
```

### Workflow 2: Create New Spreadsheet

```
1. create_sheet(file_path, "Data") - Add worksheet
2. write_cells(file_path, "Data", "A1:C1", [["Name","Age","City"]]) - Headers
3. write_cells(file_path, "Data", "A2:C4", [["Alice",25,"NY"],...]) - Data
4. write_formula(file_path, "Data", "D1", "=SUM(B2:B4)") - Calculations
```

### Workflow 3: Analyze Existing Data

```
1. get_sheet_info(file_path, sheet_name) - Understand columns
2. get_column_stats(file_path, sheet_name, "Sales") - Statistics
3. group_by(file_path, sheet_name, ["Region"], "Sales", "sum") - Grouping
4. filter_rows(file_path, sheet_name, [{"column":"Amount","operator":">","value":1000}])
```

## Error Handling

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| FileNotFoundError | Path incorrect | Verify absolute path exists |
| PermissionError | File locked | Close Excel before accessing |
| SheetNotFoundError | Wrong sheet name | Use list_sheets() first |
| ValueError | Invalid parameters | Check parameter types |

### Best Practices

1. **Always wrap in try-catch** - Handle potential errors gracefully
2. **Provide helpful error messages** - Explain what went wrong
3. **Suggest solutions** - Offer next steps to user
4. **Don't expose internals** - Keep error messages user-friendly

## Performance Tips

1. **Use pagination** - Don't read entire large files
2. **Cache results** - Reuse data when possible
3. **Batch operations** - Use write_range over multiple write_cell
4. **Minimize round trips** - Plan all operations before executing

## Security Considerations

1. **Validate file paths** - Ensure they're absolute and safe
2. **Check permissions** - Don't attempt unauthorized access
3. **Sanitize inputs** - Prevent injection attacks
4. **Don't expose secrets** - Keep credentials out of responses

## Response Format

Always format responses clearly:

```
✓ GOOD:
"Successfully read 100 rows from Sales sheet.
Data includes: Product (text), Quantity (number), Revenue (number)
Total Revenue: $125,000"

✗ BAD:
"Data retrieved."
```

## Questions to Ask Users

Before executing, clarify:

1. **File location** - "Where is the Excel file located?"
2. **Specific data** - "Which sheet and columns do you need?"
3. **Output format** - "How would you like the results presented?"
4. **Confirmation** - "Should I proceed with writing to this file?"
