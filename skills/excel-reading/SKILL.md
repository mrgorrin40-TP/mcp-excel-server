---
name: excel-reading
description: Read and extract data from Excel files (.xlsx) using MCP tools
version: 0.1.0
license: MIT
allowed_tools:
  - read_cell
  - read_range
  - get_sheet_info
  - search_cells
  - list_sheets
  - describe_workbook
---

# Excel Reading Skill

This skill enables AI agents to efficiently read and extract data from Excel files (.xlsx) using the MCP Excel Server.

## When to Use

Use this skill when the user wants to:
- Read data from an Excel file
- Extract specific cells or ranges
- Get metadata about worksheets
- Search for values in spreadsheets
- Understand the structure of a workbook

## Prerequisites

- MCP Excel Server must be running and configured
- The Excel file must exist and be accessible
- File path must be absolute (e.g., `C:/Documents/data.xlsx` on Windows or `/home/user/data.xlsx` on Linux)

## Reading Strategies

### 1. Reading Single Cells

Use `read_cell` for individual values:

```
Tool: read_cell
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "Sheet1"
  cell: "A1"
```

**Best Practices:**
- Use cell references in A1 notation (e.g., "B2", "C10")
- Sheet names are case-sensitive
- Returns the cell value as-is (string, number, date, etc.)

### 2. Reading Ranges

Use `read_range` for multiple cells:

```
Tool: read_range
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "Sheet1"
  range: "A1:C10"
```

**Best Practices:**
- Use range notation: "A1:C10" (top-left:bottom-right)
- For large datasets, use pagination with `page_size` parameter
- Default page size is 100 rows

### 3. Getting Sheet Information

Use `get_sheet_info` to understand structure:

```
Tool: get_sheet_info
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "Sheet1"
```

Returns:
- Column names and types
- Row count
- Used range
- Sample data

### 4. Searching for Values

Use `search_cells` to find data:

```
Tool: search_cells
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  query: "search term"
  sheet_name: "Sheet1"  # optional, searches all sheets if omitted
```

### 5. Listing All Sheets

Use `list_sheets` to see available worksheets:

```
Tool: list_sheets
Parameters:
  file_path: "absolute/path/to/file.xlsx"
```

### 6. Describing Workbook Structure

Use `describe_workbook` for complete overview:

```
Tool: describe_workbook
Parameters:
  file_path: "absolute/path/to/file.xlsx"
```

## Common Workflows

### Workflow 1: Explore Unknown File

1. Call `describe_workbook` to get overview
2. Call `list_sheets` to see all worksheets
3. Call `get_sheet_info` for each relevant sheet
4. Call `read_range` to examine actual data

### Workflow 2: Extract Specific Data

1. Use `search_cells` to locate the data
2. Use `read_cell` or `read_range` to extract values
3. Process and format as needed

### Workflow 3: Validate Data

1. Call `get_sheet_info` to check data types
2. Call `read_range` with pagination for large datasets
3. Check for missing values or inconsistencies

## Error Handling

| Error | Cause | Solution |
|-------|-------|----------|
| File not found | Path is incorrect | Verify absolute path exists |
| Permission denied | File is open in Excel | Close Excel before reading |
| Invalid sheet name | Sheet doesn't exist | Use `list_sheets` to see available names |
| Memory error | File too large | Use pagination with smaller `page_size` |

## Example Usage

**User Request:** "Read the sales data from C:/reports/Q1.xlsx"

**Agent Response:**
```
1. Call describe_workbook to understand structure
2. Call list_sheets to see available worksheets
3. Call get_sheet_info on the relevant sheet
4. Call read_range to extract the data
5. Present formatted results to user
```

## Tips for Effective Reading

1. **Always use absolute paths** - Relative paths won't work
2. **Check sheet names first** - Use `list_sheets` before reading
3. **Use pagination** - For files with thousands of rows
4. **Read headers first** - Understand column meanings before data
5. **Cache results** - Reuse data if multiple operations needed

## Performance Notes

- First read may be slower (file loading)
- Subsequent reads use cached data
- Large files (>10MB) may take longer
- Use `read_only` mode for better performance when not writing
