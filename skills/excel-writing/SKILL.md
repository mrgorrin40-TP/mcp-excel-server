---
name: excel-writing
description: Write, create, and modify Excel files (.xlsx) using MCP tools
version: 0.1.0
license: MIT
allowed_tools:
  - write_cells
  - write_formula
  - create_sheet
---

# Excel Writing Skill

This skill enables AI agents to create, modify, and write data to Excel files (.xlsx) using the MCP Excel Server.

## When to Use

Use this skill when the user wants to:
- Create a new Excel file
- Write data to cells or ranges
- Add formulas to worksheets
- Create new worksheets
- Update existing spreadsheets

## Prerequisites

- MCP Excel Server must be running and configured
- Write permissions to the target directory
- File path must be absolute

## Writing Strategies

### 1. Writing Single Cells

Use `write_cell` for individual values:

```
Tool: write_cell
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "Sheet1"
  cell: "A1"
  value: "Hello World"
```

**Supported Value Types:**
- Strings: `"text"`
- Numbers: `42`, `3.14`
- Booleans: `true`, `false`
- Dates: ISO format strings

### 2. Writing Ranges

Use `write_range` for multiple cells:

```
Tool: write_range
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "Sheet1"
  range: "A1:C3"
  values: [
    ["Name", "Age", "City"],
    ["Alice", 25, "New York"],
    ["Bob", 30, "London"]
  ]
```

**Best Practices:**
- Values must match range dimensions
- Use nested arrays for rows
- Headers should be in first row

### 3. Writing Formulas

Use `write_formula` for Excel formulas:

```
Tool: write_formula
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "Sheet1"
  cell: "D1"
  formula: "=SUM(A1:C1)"
```

**Formula Tips:**
- Use standard Excel formula syntax
- Formulas are recalculated on save
- Use relative references (A1) or absolute ($A$1)

### 4. Creating New Sheets

Use `create_sheet` to add worksheets:

```
Tool: create_sheet
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "NewSheet"
```

**Notes:**
- Sheet names must be unique
- Maximum 31 characters
- No special characters: /\*?[]:

## Common Workflows

### Workflow 1: Create New Spreadsheet

1. Call `create_sheet` to add worksheets
2. Call `write_range` to add headers
3. Call `write_range` to add data rows
4. Call `write_formula` for calculated columns

### Workflow 2: Update Existing Data

1. Call `read_range` to check current values
2. Call `write_cell` or `write_range` to update
3. Verify changes with `read_range`

### Workflow 3: Add Calculations

1. Call `read_range` to get raw data
2. Calculate values or formulas
3. Call `write_formula` to add results
4. Call `write_range` for summary data

## Data Validation

### Before Writing

1. **Check file exists** - Use `describe_workbook` first
2. **Verify sheet name** - Use `list_sheets` to confirm
3. **Validate data types** - Ensure compatibility with Excel
4. **Check permissions** - Ensure write access

### After Writing

1. **Read back values** - Verify successful write
2. **Check formulas** - Ensure they calculate correctly
3. **Save if needed** - Some operations require explicit save

## Error Handling

| Error | Cause | Solution |
|-------|-------|----------|
| File not found | Path is incorrect | Verify absolute path exists |
| Permission denied | File is open in Excel | Close Excel before writing |
| Invalid sheet name | Sheet doesn't exist | Use `create_sheet` first |
| Type mismatch | Value incompatible | Check data types |
| Range overflow | Writing outside bounds | Verify range dimensions |

## Example Usage

**User Request:** "Create a sales report with headers and formulas"

**Agent Response:**
```
1. Call create_sheet for "SalesReport"
2. Call write_range with headers: ["Product", "Q1", "Q2", "Total"]
3. Call write_range with data rows
4. Call write_formula for Total column: "=B2+C2"
5. Call write_formula for summary row: "=SUM(D2:D10)"
```

## Writing Best Practices

### Data Structure

1. **Headers first** - Always start with column headers
2. **Consistent types** - Keep columns homogeneous
3. **No empty rows** - Avoid gaps in data
4. **Proper formatting** - Use appropriate data types

### Performance

1. **Batch writes** - Use `write_range` over multiple `write_cell`
2. **Minimize writes** - Plan all changes before executing
3. **Avoid overwriting** - Check before modifying existing data

### Error Prevention

1. **Validate input** - Check data before writing
2. **Use try-catch** - Handle potential errors gracefully
3. **Backup data** - Keep original values for restoration

## Advanced Techniques

### Writing Formulas with Cell References

```python
# Dynamic formula based on data
for i in range(1, 11):
    formula = f"=SUM(A{i}:C{i})"
    write_formula(f"D{i}", formula)
```

### Conditional Formatting via Formulas

```python
# Add conditional formula
write_formula("E1", '=IF(D1>100,"High","Low")')
```

### Array Formulas

```python
# For Excel 365 dynamic arrays
write_formula("F1", "=UNIQUE(A1:A100)")
```

## Tips for Effective Writing

1. **Plan your structure** - Design before implementing
2. **Use meaningful headers** - Clear column names
3. **Validate before write** - Check data integrity
4. **Test with small data** - Verify approach first
5. **Document changes** - Keep track of modifications

## Performance Notes

- Small writes are fast (< 1 second)
- Large ranges may take a few seconds
- Formulas are calculated on save
- Multiple writes are batched automatically
