# Troubleshooting Guide

Common issues and solutions for MCP Excel Server.

## Table of Contents

- [Connection Issues](#connection-issues)
- [File Access Issues](#file-access-issues)
- [Performance Issues](#performance-issues)
- [Data Issues](#data-issues)
- [Configuration Issues](#configuration-issues)

---

## Connection Issues

### Server Not Responding

**Symptoms:**
- Client shows "connection refused"
- Timeout errors
- Server doesn't appear in client

**Solutions:**

1. **Check server is running:**
   ```bash
   # Start server manually to see errors
   python -m mcp_excel.server
   ```

2. **Verify Python path:**
   ```bash
   which python
   python --version  # Need 3.10+
   ```

3. **Check dependencies:**
   ```bash
   pip list | grep -E "fastmcp|openpyxl|pydantic"
   ```

4. **Verify configuration:**
   ```bash
   cat .env
   env | grep MCP_EXCEL
   ```

### Client Can't Connect

**Symptoms:**
- Client shows "unable to connect"
- MCP tools not appearing

**Solutions:**

1. **Check client configuration:**
   - Verify JSON syntax is valid
   - Ensure absolute paths
   - Check environment variables

2. **Test with MCP Inspector:**
   ```bash
   npx @modelcontextprotocol/inspector python -m mcp_excel.server
   ```

3. **Check logs:**
   ```bash
   MCP_EXCEL_LOG_LEVEL=DEBUG python -m mcp_excel.server
   ```

### Intermittent Disconnections

**Symptoms:**
- Connection drops randomly
- Tools stop responding

**Solutions:**

1. **Increase cache size:**
   ```bash
   MCP_EXCEL_CACHE_SIZE=10
   ```

2. **Check memory usage:**
   - Monitor system memory
   - Reduce cache if needed

3. **Verify file permissions:**
   - Ensure files remain accessible
   - Check for file locks

---

## File Access Issues

### File Not Found

**Symptoms:**
- `FileNotFoundError` error
- "File does not exist" message

**Solutions:**

1. **Use absolute paths:**
   ```
   ✓ CORRECT: "C:/Documents/data.xlsx"
   ✓ CORRECT: "/home/user/data.xlsx"
   ✗ WRONG: "data.xlsx"
   ✗ WRONG: "../data.xlsx"
   ```

2. **Verify file exists:**
   ```bash
   ls -la /path/to/file.xlsx
   ```

3. **Check path syntax:**
   - Windows: Use forward slashes or escaped backslashes
   - Linux: Standard Unix paths

### Permission Denied

**Symptoms:**
- `PermissionError` error
- "Access denied" message

**Solutions:**

1. **Close Excel:**
   - Excel locks files when open
   - Close Excel completely before using MCP

2. **Check file permissions:**
   ```bash
   # Linux/Mac
   chmod 644 /path/to/file.xlsx
   
   # Windows
   # Right-click → Properties → Security
   ```

3. **Run as appropriate user:**
   - Ensure server runs as user with file access

### File Locked by Another Process

**Symptoms:**
- "File is in use" error
- Cannot read or write

**Solutions:**

1. **Close all applications using the file:**
   - Excel
   - Other MCP servers
   - Backup software

2. **Check for temp files:**
   ```bash
   # Look for ~filename.xlsx
   ls -la /path/to/~*.xlsx
   ```

3. **Restart server:**
   - Kill existing server process
   - Start fresh instance

### Invalid Excel File

**Symptoms:**
- "Not a valid zip file" error
- "Corrupted file" message

**Solutions:**

1. **Verify file format:**
   - Must be .xlsx (not .xls)
   - Not corrupted or incomplete

2. **Recreate file:**
   - Open in Excel
   - Save as .xlsx again

3. **Check file size:**
   - Zero-byte files are invalid
   - Ensure complete download

---

## Performance Issues

### Slow Read Operations

**Symptoms:**
- Long wait times
- Timeout errors

**Solutions:**

1. **Use pagination:**
   ```python
   # Instead of
   read_range(range="A1:Z10000")
   
   # Use
   read_range(range="A1:Z100", page_size=100)
   ```

2. **Read specific ranges:**
   ```python
   # Instead of
   read_range(range="A1:Z1000")
   
   # Use
   read_range(range="A1:C100")  # Only needed columns
   ```

3. **Check file size:**
   - Large files (>10MB) are slower
   - Consider splitting data

### Memory Errors

**Symptoms:**
- `MemoryError` exception
- Server crashes

**Solutions:**

1. **Reduce cache size:**
   ```bash
   MCP_EXCEL_CACHE_SIZE=2
   ```

2. **Use read-only mode:**
   - Don't load unnecessary data
   - Close workbooks when done

3. **Process in chunks:**
   - Read data in pages
   - Process incrementally

### Slow Write Operations

**Symptoms:**
- Writes take too long
- Timeout on large writes

**Solutions:**

1. **Batch writes:**
   ```python
   # Instead of multiple
   write_cell("A1", "value1")
   write_cell("A2", "value2")
   
   # Use single
   write_range("A1:A2", [["value1"], ["value2"]])
   ```

2. **Minimize formula recalculations:**
   - Add formulas after data
   - Use manual calculation mode

3. **Close other applications:**
   - Free up system resources

---

## Data Issues

### Incorrect Data Types

**Symptoms:**
- Numbers read as strings
- Dates displayed incorrectly

**Solutions:**

1. **Check Excel formatting:**
   - Ensure cells have correct format
   - Text cells should be text, not numbers

2. **Use data_only parameter:**
   ```python
   # Read computed values
   read_cell(cell="A1", data_only=True)
   ```

3. **Convert in code:**
   ```python
   # After reading
   value = int(cell_value) if cell_value else 0
   ```

### Missing Data

**Symptoms:**
- Empty cells where data expected
- Incomplete results

**Solutions:**

1. **Check for merged cells:**
   - Merged cells only have value in top-left
   - Unmerge if needed

2. **Verify used range:**
   ```python
   get_sheet_info()  # Check row/column counts
   ```

3. **Check for hidden rows/columns:**
   - Hidden data may not be read

### Formulas Not Working

**Symptoms:**
- Formulas show as text
- Calculations not updating

**Solutions:**

1. **Check formula syntax:**
   - Must start with `=`
   - Use correct function names

2. **Verify cell references:**
   - Check row/column numbers
   - Use absolute references if needed

3. **Recalculate:**
   - Open in Excel and save
   - Force recalculation

### Search Not Finding Data

**Symptoms:**
- `search_cells` returns empty results
- Data exists but not found

**Solutions:**

1. **Check search term:**
   - Case-sensitive by default
   - Check for extra spaces

2. **Verify sheet name:**
   - Must match exactly
   - Use `list_sheets()` to confirm

3. **Search all sheets:**
   - Omit `sheet_name` parameter

---

## Configuration Issues

### Environment Variables Not Loading

**Symptoms:**
- Default values used instead of custom
- Settings not applied

**Solutions:**

1. **Check .env file location:**
   - Must be in project root
   - Check file permissions

2. **Verify .env syntax:**
   ```
   ✓ CORRECT: MCP_EXCEL_LOG_LEVEL=DEBUG
   ✗ WRONG: MCP_EXCEL_LOG_LEVEL = DEBUG
   ✗ WRONG: "MCP_EXCEL_LOG_LEVEL"="DEBUG"
   ```

3. **Restart server:**
   - Changes require restart

### Wrong Python Version

**Symptoms:**
- Syntax errors
- Import errors

**Solutions:**

1. **Check version:**
   ```bash
   python --version
   ```

2. **Use correct Python:**
   ```bash
   python3 --version  # Linux/Mac
   py --version       # Windows
   ```

3. **Update Python:**
   - Download from python.org
   - Install 3.10 or higher

### Missing Dependencies

**Symptoms:**
- `ModuleNotFoundError`
- Import failures

**Solutions:**

1. **Install dependencies:**
   ```bash
   pip install -e .
   # or
   poetry install
   ```

2. **Check installed packages:**
   ```bash
   pip list | grep -E "fastmcp|openpyxl|pydantic"
   ```

3. **Use virtual environment:**
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Linux/Mac
   .venv\Scripts\activate     # Windows
   pip install -e .
   ```

---

## Getting Help

### Enable Debug Logging

```bash
MCP_EXCEL_LOG_LEVEL=DEBUG MCP_EXCEL_MASK_ERRORS=false python -m mcp_excel.server
```

### Check Server Status

```bash
# Test with MCP Inspector
npx @modelcontextprotocol/inspector python -m mcp_excel.server
```

### Report Issues

When reporting issues, include:
1. Error message (full traceback)
2. Steps to reproduce
3. Environment (OS, Python version)
4. Configuration (.env contents)
5. File details (size, format)

### Community Resources

- [GitHub Issues](https://github.com/username/mcp-excel-server/issues)
- [MCP Documentation](https://modelcontextprotocol.io)
- [FastMCP Documentation](https://gofastmcp.com)
