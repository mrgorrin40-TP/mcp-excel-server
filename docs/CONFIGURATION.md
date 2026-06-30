# Configuration Guide

Complete guide to configuring MCP Excel Server.

## Table of Contents

- [Environment Variables](#environment-variables)
- [Transport Configuration](#transport-configuration)
- [Cache Settings](#cache-settings)
- [Response Limits](#response-limits)
- [Logging](#logging)
- [Client Configuration](#client-configuration)

---

## Environment Variables

All configuration is done via environment variables with the prefix `MCP_EXCEL_`.

### Copy Environment Template

```bash
cp .env.example .env
```

### Available Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `MCP_EXCEL_TRANSPORT` | `stdio` | Transport type: `stdio` or `http` |
| `MCP_EXCEL_MASK_ERRORS` | `true` | Mask internal errors in production |
| `MCP_EXCEL_CACHE_SIZE` | `5` | Maximum workbooks in cache |
| `MCP_EXCEL_CACHE_MEMORY_MB` | `1024` | Maximum cache memory (MB) |
| `MCP_EXCEL_PAGING_LIMIT` | `4000` | Maximum cells per page |
| `MCP_EXCEL_MAX_ROWS` | `500` | Maximum rows in response |
| `MCP_EXCEL_MAX_COLUMNS` | `50` | Maximum columns in response |
| `MCP_EXCEL_LOG_LEVEL` | `INFO` | Logging level |

---

## Transport Configuration

### stdio (Default)

Best for local development and personal use.

```bash
MCP_EXCEL_TRANSPORT=stdio
```

**Pros:**
- No network configuration needed
- Low latency
- Simple setup

**Cons:**
- Single user only
- No remote access

### HTTP

Best for teams, remote access, and production deployments.

```bash
MCP_EXCEL_TRANSPORT=http
MCP_EXCEL_HTTP_HOST=0.0.0.0
MCP_EXCEL_HTTP_PORT=8000
```

**Pros:**
- Multi-user support
- Remote access
- Scalable

**Cons:**
- Requires network configuration
- Higher latency
- More complex setup

---

## Cache Settings

The server caches loaded workbooks in memory for better performance.

### Cache Size

```bash
MCP_EXCEL_CACHE_SIZE=5
```

- Maximum number of workbooks to keep in memory
- LRU eviction when limit is reached
- Increase for multi-file workflows
- Decrease for limited memory

### Cache Memory Limit

```bash
MCP_EXCEL_CACHE_MEMORY_MB=1024
```

- Maximum memory for cached workbooks (MB)
- Older workbooks evicted when exceeded
- Monitor with server logs

### Cache Behavior

1. **First access**: File loaded from disk
2. **Subsequent access**: Served from cache
3. **Write operations**: Cache updated immediately
4. **Eviction**: Least recently used removed

---

## Response Limits

Control response sizes to prevent context overflow.

### Maximum Rows

```bash
MCP_EXCEL_MAX_ROWS=500
```

- Maximum rows returned in single response
- Use pagination for more data
- Reduces token usage

### Maximum Columns

```bash
MCP_EXCEL_MAX_COLUMNS=50
```

- Maximum columns returned in response
- Prevents wide table overflow
- Adjust based on your data

### Pagination Limit

```bash
MCP_EXCEL_PAGING_LIMIT=4000
```

- Maximum cells per page
- Used for automatic pagination
- Larger values = more data per request

---

## Logging

### Log Level

```bash
MCP_EXCEL_LOG_LEVEL=INFO
```

**Available levels:**

| Level | Description |
|-------|-------------|
| `DEBUG` | Detailed debugging information |
| `INFO` | General information |
| `WARNING` | Warning messages |
| `ERROR` | Error messages only |

### Log Output

By default, logs go to stderr (not stdout):

```python
# In server.py
import logging
logging.basicConfig(
    level=settings.log_level,
    stream=sys.stderr  # Keep stdout clean for MCP
)
```

### Debug Mode

For development, enable debug logging:

```bash
MCP_EXCEL_LOG_LEVEL=DEBUG
MCP_EXCEL_MASK_ERRORS=false
```

---

## Client Configuration

### Claude Desktop

**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.server"],
      "cwd": "/path/to/mcp-excel-server",
      "env": {
        "MCP_EXCEL_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

### VS Code

`.vscode/mcp.json` in workspace root:

```json
{
  "servers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.server"],
      "cwd": "${workspaceFolder}"
    }
  }
}
```

### Cursor

`.cursor/mcp.json` in workspace root:

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

Add to OpenCode configuration:

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "mcp_excel.server"],
      "env": {
        "MCP_EXCEL_TRANSPORT": "stdio",
        "MCP_EXCEL_LOG_LEVEL": "INFO"
      }
    }
  }
}
```

### Custom Python Path

If Python is not in PATH:

```json
{
  "mcpServers": {
    "excel": {
      "command": "/usr/bin/python3",
      "args": ["-m", "mcp_excel.server"],
      "cwd": "/path/to/mcp-excel-server"
    }
  }
}
```

### Using Poetry

```json
{
  "mcpServers": {
    "excel": {
      "command": "poetry",
      "args": ["run", "python", "-m", "mcp_excel.server"],
      "cwd": "/path/to/mcp-excel-server"
    }
  }
}
```

### Using Virtual Environment

```json
{
  "mcpServers": {
    "excel": {
      "command": "/path/to/mcp-excel-server/.venv/bin/python",
      "args": ["-m", "mcp_excel.server"],
      "cwd": "/path/to/mcp-excel-server"
    }
  }
}
```

---

## Advanced Configuration

### Custom Cache Directory

```python
# In config.py
cache_dir: Path = Field(default=Path("~/.mcp-excel/cache"))
```

### Timeout Settings

```python
# In config.py
file_timeout: int = Field(default=30, description="File operation timeout in seconds")
```

### Security Settings

```python
# In config.py
allowed_directories: list[str] = Field(
    default=["/home", "C:/Users"],
    description="Directories allowed for file access"
)
```

---

## Environment-Specific Configs

### Development

```bash
MCP_EXCEL_TRANSPORT=stdio
MCP_EXCEL_LOG_LEVEL=DEBUG
MCP_EXCEL_MASK_ERRORS=false
MCP_EXCEL_CACHE_SIZE=2
```

### Production

```bash
MCP_EXCEL_TRANSPORT=http
MCP_EXCEL_LOG_LEVEL=WARNING
MCP_EXCEL_MASK_ERRORS=true
MCP_EXCEL_CACHE_SIZE=10
MCP_EXCEL_CACHE_MEMORY_MB=2048
```

### Testing

```bash
MCP_EXCEL_TRANSPORT=stdio
MCP_EXCEL_LOG_LEVEL=DEBUG
MCP_EXCEL_CACHE_SIZE=1
MCP_EXCEL_MAX_ROWS=100
```

---

## Troubleshooting Configuration

### Server Won't Start

1. Check Python version: `python --version` (need 3.10+)
2. Verify dependencies: `pip list | grep -E "fastmcp|openpyxl"`
3. Check environment: `env | grep MCP_EXCEL`

### Configuration Not Loading

1. Verify `.env` file exists in project root
2. Check file permissions
3. Ensure no syntax errors in `.env`

### Wrong Settings Applied

1. Check for conflicting environment variables
2. Verify `.env` file location
3. Restart server after config changes
