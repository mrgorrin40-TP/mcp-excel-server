# MCP Excel Server

[![MCP Compatible](https://img.shields.io/badge/MCP-Compatible-green.svg)](https://modelcontextprotocol.io)
[![Python 3.10+](https://img.shields.io/badge/Python-3.10+-blue.svg)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

A Model Context Protocol (MCP) server for Excel file manipulation. Read, write, analyze, and transform Excel spreadsheets (.xlsx) using AI agents.

## Features

- **Read Data** - Extract values, ranges, and metadata from Excel files
- **Write Data** - Create and modify spreadsheets programmatically
- **Formulas** - Add Excel formulas for calculations
- **Charts** - Create and modify Excel charts (bar, line, pie, scatter, area)
- **Tables** - Work with Excel structured tables
- **Analysis** - Statistical summaries, filtering, and grouping

## Quick Start

### Installation

```bash
# Clone the repository
git clone https://github.com/username/mcp-excel-server.git
cd mcp-excel-server

# Install with Poetry
poetry install

# Or with pip
pip install -e .
```

### Configuration

```bash
# Copy environment template
cp .env.example .env

# Edit .env with your settings
# MCP_EXCEL_TRANSPORT=stdio
# MCP_EXCEL_CACHE_SIZE=5
```

### Running the Server

```bash
# Using Poetry
poetry run python -m mcp_excel.server

# Or directly
python -m mcp_excel.server
```

## Client Configuration

### Claude Desktop

Add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.server"],
      "cwd": "/path/to/mcp-excel-server"
    }
  }
}
```

### VS Code / Cursor

Add to `.vscode/mcp.json` or `.cursor/mcp.json`:

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.server"],
      "cwd": "${workspaceFolder}"
    }
  }
}
```

### OpenCode

Add to your OpenCode configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.server"],
      "env": {
        "MCP_EXCEL_TRANSPORT": "stdio"
      }
    }
  }
}
```

## Available Tools

### Reading Data

| Tool | Description |
|------|-------------|
| `read_cell` | Read a single cell value |
| `read_range` | Read a range of cells |
| `get_sheet_info` | Get worksheet metadata |
| `search_cells` | Search for values |
| `list_sheets` | List all worksheets |
| `describe_workbook` | Get workbook overview |

### Writing Data

| Tool | Description |
|------|-------------|
| `write_cells` | Write values to cells |
| `write_formula` | Add Excel formulas |
| `create_sheet` | Create new worksheet |
| `delete_sheet` | Delete a worksheet |

### Formulas

| Tool | Description |
|------|-------------|
| `read_formula` | Read formula text from a cell |
| `get_formula_templates` | Get common formula templates |

### Analysis

| Tool | Description |
|------|-------------|
| `get_column_stats` | Statistical summary |
| `filter_rows` | Filter data by conditions |
| `group_by` | Group and aggregate data |

### Charts

| Tool | Description |
|------|-------------|
| `create_chart` | Create bar, line, pie, scatter, or area chart |
| `list_charts` | List all charts in a worksheet |
| `modify_chart` | Update chart properties |
| `delete_chart` | Remove a chart |

### Tables

| Tool | Description |
|------|-------------|
| `create_table` | Create Excel structured table |
| `list_tables` | List all tables in a worksheet |
| `delete_table` | Delete a table |
| `add_table_row` | Add row to a table |

## Example Usage

### Read Sales Data

```
User: Read the sales data from C:/reports/Q1.xlsx

Agent calls:
1. describe_workbook(file_path="C:/reports/Q1.xlsx")
2. list_sheets(file_path="C:/reports/Q1.xlsx")
3. read_range(file_path="C:/reports/Q1.xlsx", sheet_name="Sales", range="A1:F100")
```

### Create Summary Report

```
User: Create a monthly summary with totals

Agent calls:
1. create_sheet(file_path="report.xlsx", sheet_name="Summary")
2. write_cells(file_path="report.xlsx", sheet_name="Summary", range="A1:D1", values=[["Month", "Sales", "Cost", "Profit"]])
3. write_formula(file_path="report.xlsx", sheet_name="Summary", cell="D2", formula="=B2-C2")
```

### Analyze Data

```
User: Analyze customer demographics

Agent calls:
1. get_column_stats(file_path="customers.xlsx", sheet_name="Data", column="Age")
2. group_by(file_path="customers.xlsx", sheet_name="Data", columns=["Region"], agg_column="Revenue", agg_func="sum")
3. filter_rows(file_path="customers.xlsx", sheet_name="Data", filters=[{"column": "Status", "operator": "==", "value": "Active"}])
```

## Skills

This server includes agent skills for better interaction:

- **excel-reading** - Guide for reading Excel files
- **excel-writing** - Guide for writing to Excel files
- **excel-analysis** - Guide for data analysis
- **excel-formulas** - Guide for Excel formulas

See the `skills/` directory for detailed documentation.

## Development

### Setup

```bash
# Install development dependencies
poetry install --with dev

# Run tests
poetry run pytest

# Run linting
poetry run ruff check src/

# Run type checking
poetry run mypy src/
```

### Project Structure

```
mcp-excel-server/
├── src/mcp_excel/           # Source code
│   ├── server.py           # MCP server entry point
│   ├── config.py           # Configuration
│   ├── backends/           # Excel backends
│   ├── tools/              # MCP tools
│   ├── resources/          # MCP resources
│   ├── prompts/            # MCP prompts
│   └── utils/              # Utilities
├── tests/                  # Test suite
├── skills/                 # Agent skills
├── examples/               # Usage examples
└── docs/                   # Documentation
```

### Adding New Tools

1. Create tool in `src/mcp_excel/tools/`
2. Add Pydantic models for input/output
3. Register in `server.py`
4. Add tests in `tests/`
5. Update documentation

## Troubleshooting

### Common Issues

**File not found error:**
- Ensure you're using absolute paths
- Check file exists and is accessible

**Permission denied:**
- Close Excel if file is open
- Check file permissions

**Memory errors:**
- Use pagination for large files
- Reduce cache size in `.env`

See `docs/TROUBLESHOOTING.md` for more details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests
5. Run the test suite
6. Submit a pull request

## License

MIT License - see [LICENSE](LICENSE) for details.

## Acknowledgments

- [Model Context Protocol](https://modelcontextprotocol.io) - The standard for AI-tool integration
- [FastMCP](https://github.com/PrefectHQ/fastmcp) - Python MCP framework
- [openpyxl](https://openpyxl.readthedocs.io) - Excel file handling
