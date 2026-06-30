"""Basic usage example for MCP Excel Server."""

# This example demonstrates how to use the MCP Excel Server tools
# from within an AI agent context.

# Example 1: Read data from an Excel file
example_read = """
User: Read the sales data from C:/reports/Q1.xlsx

Agent calls:
1. describe_workbook(file_path="C:/reports/Q1.xlsx")
   → Gets overview of the workbook structure

2. list_sheets(file_path="C:/reports/Q1.xlsx")
   → Sees available worksheets: ["Sales", "Summary", "Charts"]

3. get_sheet_info(file_path="C:/reports/Q1.xlsx", sheet_name="Sales")
   → Gets column names, types, and sample data

4. read_range(file_path="C:/reports/Q1.xlsx", sheet_name="Sales", range="A1:F100")
   → Reads the actual data
"""

# Example 2: Create a new spreadsheet
example_create = """
User: Create a monthly budget spreadsheet

Agent calls:
1. create_sheet(file_path="budget.xlsx", sheet_name="Budget")
   → Creates a new worksheet

2. write_cells(
     file_path="budget.xlsx",
     sheet_name="Budget",
     range="A1:D1",
     values=[["Category", "Budget", "Actual", "Difference"]]
   )
   → Adds column headers

3. write_cells(
     file_path="budget.xlsx",
     sheet_name="Budget",
     range="A2:D5",
     values=[
       ["Rent", 1500, 1500, "=B2-C2"],
       ["Food", 500, 450, "=B3-C3"],
       ["Transport", 200, 180, "=B4-C4"]
     ]
   )
   → Adds data with formulas

4. write_formula(
     file_path="budget.xlsx",
     sheet_name="Budget",
     cell="D6",
     formula="=SUM(D2:D4)"
   )
   → Adds total formula
"""

# Example 3: Analyze data
example_analyze = """
User: Analyze customer demographics

Agent calls:
1. get_sheet_info(file_path="customers.xlsx", sheet_name="Data")
   → Understands the data structure

2. get_column_stats(file_path="customers.xlsx", sheet_name="Data", column="Age")
   → Gets age statistics: mean=35, median=33, min=18, max=75

3. group_by(
     file_path="customers.xlsx",
     sheet_name="Data",
     columns=["Region"],
     agg_column="Revenue",
     agg_func="sum"
   )
   → Gets revenue by region

4. filter_rows(
     file_path="customers.xlsx",
     sheet_name="Data",
     filters=[{"column": "Status", "operator": "==", "value": "Active"}]
   )
   → Gets only active customers
"""

# Example 4: Search for data
example_search = """
User: Find all orders from New York

Agent calls:
1. search_cells(
     file_path="orders.xlsx",
     query="New York"
   )
   → Finds all cells containing "New York"

2. For each match, reads the full row to get order details
"""

# Example 5: Update existing data
example_update = """
User: Update the price of Product A to $29.99

Agent calls:
1. search_cells(
     file_path="products.xlsx",
     query="Product A"
   )
   → Finds the cell containing "Product A"

2. write_cells(
     file_path="products.xlsx",
     sheet_name="Products",
     range="C5",
     values=[[29.99]]
   )
   → Updates the price
"""

if __name__ == "__main__":
    print("MCP Excel Server - Usage Examples")
    print("=" * 50)
    print("\nExample 1: Read Data")
    print(example_read)
    print("\nExample 2: Create Spreadsheet")
    print(example_create)
    print("\nExample 3: Analyze Data")
    print(example_analyze)
    print("\nExample 4: Search Data")
    print(example_search)
    print("\nExample 5: Update Data")
    print(example_update)
