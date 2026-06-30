---
name: excel-analysis
description: Analyze Excel data with statistics, filtering, and grouping operations
version: 0.1.0
license: MIT
allowed_tools:
  - get_column_stats
  - filter_rows
  - group_by
  - read_range
  - get_sheet_info
---

# Excel Analysis Skill

This skill enables AI agents to perform data analysis on Excel files using statistical operations, filtering, and grouping.

## When to Use

Use this skill when the user wants to:
- Get statistical summaries of data
- Filter rows based on conditions
- Group data by columns and aggregate
- Identify patterns or anomalies
- Generate insights from spreadsheets

## Prerequisites

- MCP Excel Server must be running
- Data must be in tabular format with headers
- File path must be absolute

## Analysis Strategies

### 1. Column Statistics

Use `get_column_stats` for statistical summaries:

```
Tool: get_column_stats
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "Sheet1"
  column: "B"
```

**Returns:**
- Count, sum, mean, median
- Min, max, range
- Standard deviation, variance
- Null count
- Data type information

### 2. Filtering Rows

Use `filter_rows` to subset data:

```
Tool: filter_rows
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "Sheet1"
  filters: [
    {"column": "B", "operator": ">", "value": 100},
    {"column": "C", "operator": "==", "value": "Active"}
  ]
  logic: "AND"
```

**Operators:**
- Comparison: `==`, `!=`, `>`, `<`, `>=`, `<=`
- Text: `contains`, `startswith`, `endswith`
- Set: `in`, `not_in`
- Null: `is_null`, `is_not_null`

### 3. Grouping and Aggregation

Use `group_by` for pivot-like analysis:

```
Tool: group_by
Parameters:
  file_path: "absolute/path/to/file.xlsx"
  sheet_name: "Sheet1"
  columns: ["Category", "Region"]
  agg_column: "Sales"
  agg_func: "sum"
```

**Aggregation Functions:**
- `sum` - Total
- `mean` - Average
- `median` - Middle value
- `count` - Count of rows
- `min` - Minimum
- `max` - Maximum
- `std` - Standard deviation

## Common Workflows

### Workflow 1: Data Exploration

1. Call `get_sheet_info` to understand structure
2. Call `read_range` to examine sample data
3. Call `get_column_stats` for each important column
4. Identify key patterns and insights

### Workflow 2: Data Validation

1. Call `get_column_stats` to check distributions
2. Call `filter_rows` to find outliers
3. Call `get_column_stats` on filtered data
4. Document any anomalies found

### Workflow 3: Report Generation

1. Call `get_sheet_info` for overview
2. Call `get_column_stats` for key metrics
3. Call `group_by` for categorical breakdowns
4. Call `filter_rows` for specific segments
5. Compile findings into report

## Analysis Patterns

### Pattern 1: Summary Statistics

```python
# Get overview of all numeric columns
for column in ["Sales", "Quantity", "Price"]:
    stats = get_column_stats(column)
    print(f"{column}: mean={stats['mean']}, std={stats['std']}")
```

### Pattern 2: Conditional Analysis

```python
# Analyze high-value transactions
high_value = filter_rows([
    {"column": "Amount", "operator": ">", "value": 1000}
])
# Then get stats on filtered data
```

### Pattern 3: Cross-tabulation

```python
# Group by multiple dimensions
group_by(
    columns=["Category", "Month"],
    agg_column="Revenue",
    agg_func="sum"
)
```

## Interpreting Results

### Statistical Measures

| Measure | Meaning | Use Case |
|---------|---------|----------|
| Mean | Average value | Central tendency |
| Median | Middle value | Skewed data |
| Std Dev | Spread of data | Consistency |
| Min/Max | Range | Boundaries |
| Count | Number of records | Volume |

### Filter Results

- Returns matching rows
- Preserves original order
- Includes all columns
- May be paginated for large results

### Group Results

- One row per group
- Aggregated values
- Sorted by group columns
- Includes group labels

## Common Analysis Tasks

### Sales Analysis

```
1. get_column_stats on "Sales" column
2. group_by "Region" with sum of "Sales"
3. group_by "Product" with mean of "Sales"
4. filter_rows where "Sales" > target
```

### Data Quality Check

```
1. get_column_stats on each column
2. filter_rows where columns are null
3. get_column_stats on filtered data
4. Document quality issues
```

### Trend Analysis

```
1. group_by "Month" with sum of "Revenue"
2. get_column_stats on "Revenue"
3. Identify peaks and valleys
4. Calculate growth rates
```

## Error Handling

| Error | Cause | Solution |
|-------|-------|----------|
| Column not found | Invalid column name | Use `get_sheet_info` to see columns |
| Invalid operator | Typo in operator | Check operator spelling |
| Type mismatch | Comparing incompatible types | Verify column data type |
| Memory error | Too many results | Use more specific filters |

## Best Practices

### Data Preparation

1. **Clean data first** - Remove empty rows/columns
2. **Consistent formatting** - Ensure uniform data types
3. **Clear headers** - Use descriptive column names
4. **No merged cells** - Avoid complex cell structures

### Analysis Design

1. **Start broad** - Get overview first
2. **Narrow focus** - Use filters to drill down
3. **Validate findings** - Cross-check with different methods
4. **Document assumptions** - Note any data limitations

### Result Interpretation

1. **Check sample size** - Small samples may be misleading
2. **Consider context** - Business meaning matters
3. **Look for outliers** - They may indicate errors
4. **Compare to baseline** - Historical or industry data

## Example Analysis

**User Request:** "Analyze our Q1 sales performance"

**Agent Response:**
```
1. Call describe_workbook for overview
2. Call get_column_stats on "Sales" column
3. Call group_by on "Region" with sum of "Sales"
4. Call group_by on "Product" with mean of "Sales"
5. Call filter_rows where "Sales" > 10000
6. Present summary with key insights:
   - Total sales: $X
   - Top region: Y
   - Best product: Z
   - High-value transactions: N
```

## Tips for Effective Analysis

1. **Understand your data** - Know what each column means
2. **Ask clarifying questions** - What specifically to analyze?
3. **Use appropriate methods** - Match technique to question
4. **Validate results** - Sanity check findings
5. **Communicate clearly** - Present insights simply
