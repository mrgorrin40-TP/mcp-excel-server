---
name: excel-formulas
description: Create and manage Excel formulas for calculations and data transformations
version: 0.1.0
license: MIT
allowed_tools:
  - write_formula
  - read_cell
  - read_range
  - write_cells
---

# Excel Formulas Skill

This skill enables AI agents to create, understand, and manage Excel formulas for calculations and data transformations.

## When to Use

Use this skill when the user wants to:
- Add calculations to spreadsheets
- Create dynamic formulas
- Build summary reports
- Perform data transformations
- Understand existing formulas

## Prerequisites

- MCP Excel Server must be running
- Understanding of Excel formula syntax
- Data must be properly structured

## Formula Fundamentals

### Basic Syntax

Excel formulas start with `=`:

```
=SUM(A1:A10)
=AVERAGE(B1:B100)
=IF(C1>100,"High","Low")
```

### Cell References

| Type | Example | Meaning |
|------|---------|---------|
| Relative | A1 | Changes when copied |
| Absolute | $A$1 | Never changes |
| Mixed | $A1 | Column fixed, row changes |
| Mixed | A$1 | Row fixed, column changes |

### Common Functions

#### Math & Statistics
```
=SUM(range)           # Add values
=AVERAGE(range)       # Mean
=MEDIAN(range)        # Middle value
=MIN(range)           # Minimum
=MAX(range)           # Maximum
=COUNT(range)         # Count numbers
=COUNTA(range)        # Count non-empty
=STDEV(range)         # Standard deviation
```

#### Logic
```
=IF(condition, true, false)        # Conditional
=AND(condition1, condition2)       # All true
=OR(condition1, condition2)        # Any true
=NOT(condition)                    # Reverse
=IFS(cond1,val1,cond2,val2,...)   # Multiple conditions
```

#### Text
```
=CONCATENATE(text1, text2)   # Join text
=LEFT(text, num)              # Left characters
=RIGHT(text, num)             # Right characters
=LEN(text)                    # Length
=UPPER(text)                  # Uppercase
=LOWER(text)                  # Lowercase
=TRIM(text)                   # Remove spaces
```

#### Lookup
```
=VLOOKUP(value, range, col, match)   # Vertical lookup
=HLOOKUP(value, range, row, match)   # Horizontal lookup
=INDEX(range, row, col)              # Return value
=MATCH(value, range, type)           # Find position
```

## Writing Formulas

### Basic Formula Writing

```
Tool: write_formula
Parameters:
  file_path: "path/to/file.xlsx"
  sheet_name: "Sheet1"
  cell: "D1"
  formula: "=SUM(A1:C1)"
```

### Formula Patterns

#### Running Total
```excel
=SUM($A$1:A1)
```

#### Percentage of Total
```excel
=A1/SUM($A$1:$A$100)
```

#### Conditional Sum
```excel
=SUMIF(B:B,"Sales",C:C)
```

#### Lookup with Error Handling
```excel
=IFERROR(VLOOKUP(A1,Data!A:B,2,FALSE),"Not Found")
```

## Common Workflows

### Workflow 1: Add Summary Row

1. Call `read_range` to get data
2. Calculate totals or averages
3. Call `write_formula` for each summary cell
4. Format as needed

### Workflow 2: Create Calculated Column

1. Call `get_sheet_info` to understand data
2. Design formula based on requirements
3. Call `write_formula` for first cell
4. Copy formula to other rows (manually or via tool)

### Workflow 3: Build Dashboard

1. Create summary worksheets
2. Add formulas for key metrics
3. Link to source data
4. Create charts from formula results

## Formula Examples

### Sales Calculations

```excel
# Total Sales
=SUM(D2:D100)

# Average Sale
=AVERAGE(D2:D100)

# Sales Tax
=D2*0.08

# Discounted Price
=D2*(1-E2)

# Year-over-Year Growth
=(D2-D1)/D1
```

### Data Analysis

```excel
# Count by Category
=COUNTIF(A:A,"Electronics")

# Sum by Region
=SUMIF(B:B,"North",C:C)

# Average by Month
=AVERAGEIFS(D:D,A:A,">="&DATE(2024,1,1),A:A,"<"&DATE(2024,2,1))

# Top 10 Values
=LARGE(D2:D100,ROW()-1)
```

### Text Manipulation

```excel
# Extract First Name
=LEFT(A1,FIND(" ",A1)-1)

# Extract Last Name
=MID(A1,FIND(" ",A1)+1,LEN(A1))

# Format Phone Number
="(" & LEFT(A1,3) & ") " & MID(A1,4,3) & "-" & RIGHT(A1,4)

# Clean Text
=TRIM(CLEAN(A1))
```

### Date Operations

```excel
# Today's Date
=TODAY()

# Days Between
=B1-A1

# Month Name
=TEXT(A1,"mmmm")

# Last Day of Month
=EOMONTH(A1,0)
```

## Advanced Formulas

### Array Formulas

```excel
# Sum of Top 3 Values
=SUM(LARGE(A1:A100,{1,2,3}))

# Count Unique Values
=SUMPRODUCT(1/COUNTIF(A1:A100,A1:A100))

# Conditional Array
=SUM((A1:A100="Sales")*(B1:B100>1000)*C1:C100)
```

### Dynamic Arrays (Excel 365)

```excel
# Unique Values
=UNIQUE(A1:A100)

# Sorted List
=SORT(A1:B100,2,-1)

# Filtered Data
=FILTER(A1:C100,B1:B100>1000)

# Spill Range
=SEQUENCE(10,1,1,1)
```

## Formula Debugging

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| #REF! | Invalid reference | Check cell references |
| #VALUE! | Wrong data type | Verify argument types |
| #N/A | Value not found | Use IFERROR or check lookup |
| #DIV/0! | Division by zero | Add zero check |
| #NAME? | Unknown function | Check function spelling |
| #NULL! | Intersection error | Check range operators |

### Debugging Steps

1. **Check syntax** - Verify formula structure
2. **Verify references** - Ensure cells exist
3. **Check data types** - Numbers vs text
4. **Test in Excel** - Build formula incrementally
5. **Use Evaluate** - Step through calculation

## Best Practices

### Formula Design

1. **Keep it simple** - Break complex formulas into steps
2. **Use named ranges** - For clarity and maintenance
3. **Document assumptions** - Add comments if needed
4. **Test thoroughly** - Verify with sample data

### Performance

1. **Avoid entire column references** - Use specific ranges
2. **Minimize volatile functions** - NOW(), TODAY(), RAND()
3. **Use helper columns** - For complex calculations
4. **Consider Power Query** - For very large datasets

### Maintenance

1. **Consistent naming** - Use clear column headers
2. **Modular design** - Separate calculation areas
3. **Version control** - Track formula changes
4. **Documentation** - Explain complex logic

## Example Usage

**User Request:** "Add a column calculating profit margin"

**Agent Response:**
```
1. Call get_sheet_info to understand structure
2. Identify columns: Revenue (D), Cost (E)
3. Call write_formula for first cell:
   cell: "F2"
   formula: "=(D2-E2)/D2"
4. Format as percentage
5. Copy formula to remaining rows
```

## Tips for Formula Mastery

1. **Learn incrementally** - Master basics first
2. **Practice regularly** - Build real spreadsheets
3. **Use Excel help** - Built-in documentation
4. **Join communities** - Excel forums and groups
5. **Explore templates** - Learn from examples
