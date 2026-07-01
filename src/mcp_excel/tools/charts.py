"""MCP tools for Excel charts."""

import logging
from typing import Annotated, Any

from fastmcp import FastMCP
from pydantic import Field

from ..utils.backend import get_backend

logger = logging.getLogger(__name__)

# Create tools router
tools = FastMCP("Excel Chart Tools", mask_error_details=True)

# Chart type mapping
CHART_TYPES = {
    "bar": "bar",
    "line": "line",
    "pie": "pie",
    "scatter": "scatter",
    "area": "area",
    "bar3d": "bar3d",
    "line3d": "line3d",
    "pie3d": "pie3d",
}


@tools.tool(
    name="create_chart",
    description="Create a chart in an Excel worksheet",
    tags={"excel", "write", "chart"},
)
async def create_chart(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
    chart_type: Annotated[
        str, Field(description="Chart type: bar, line, pie, scatter, area")
    ],
    title: Annotated[str, Field(description="Chart title")],
    data_range: Annotated[str, Field(description="Data range (e.g., 'A1:D10')")],
    category_range: Annotated[
        str | None, Field(description="Category labels range (e.g., 'A1:A10')")
    ] = None,
    series_names: Annotated[
        list[str] | None, Field(description="Series names (optional)")
    ] = None,
    position: Annotated[
        str | None, Field(description="Position (e.g., 'E2', or 'E2:J20')")
    ] = None,
) -> dict[str, Any]:
    """Create a chart in an Excel worksheet."""
    try:
        backend = get_backend(file_path)

        ws = backend.get_sheet(sheet_name)

        # Import chart classes
        from openpyxl.chart import (
            AreaChart,
            BarChart,
            LineChart,
            PieChart,
            Reference,
            ScatterChart,
        )

        # Create chart based on type
        chart_type_lower = chart_type.lower()
        if chart_type_lower == "bar":
            chart = BarChart()
        elif chart_type_lower == "line":
            chart = LineChart()
        elif chart_type_lower == "pie":
            chart = PieChart()
        elif chart_type_lower == "scatter":
            chart = ScatterChart()
        elif chart_type_lower == "area":
            chart = AreaChart()
        else:
            return {"success": False, "error": f"Unsupported chart type: {chart_type}"}

        # Set title
        chart.title = title
        chart.style = 10

        # Parse data range
        data_ref = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=1)
        if ":" in data_range:
            start, end = data_range.split(":")
            start_col = _col_letter_to_index(start)
            start_row = int("".join(c for c in start if c.isdigit()))
            end_col = _col_letter_to_index(end)
            end_row = int("".join(c for c in end if c.isdigit()))
            data_ref = Reference(
                ws,
                min_col=start_col,
                min_row=start_row,
                max_col=end_col,
                max_row=end_row,
            )

        # Add data
        chart.add_data(data_ref, titles_from_data=True)

        # Add category range if specified
        if category_range:
            cat_ref = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=1)
            if ":" in category_range:
                cat_start, cat_end = category_range.split(":")
                cat_col = _col_letter_to_index(cat_start)
                cat_start_row = int("".join(c for c in cat_start if c.isdigit()))
                cat_end_row = int("".join(c for c in cat_end if c.isdigit()))
                cat_ref = Reference(
                    ws,
                    min_col=cat_col,
                    min_row=cat_start_row,
                    max_row=cat_end_row,
                )
            chart.set_categories(cat_ref)

        # Set position
        if position:
            if ":" in position:
                chart.anchor = position
            else:
                chart.anchor = position
        else:
            chart.anchor = "E2"

        # Add chart to worksheet
        ws.add_chart(chart, chart.anchor or "E2")

        backend.save()

        return {
            "success": True,
            "chart_type": chart_type,
            "title": title,
            "sheet_name": sheet_name,
            "position": chart.anchor,
            "data_range": data_range,
        }
    except Exception as e:
        logger.error("Error creating chart: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="list_charts",
    description="List all charts in an Excel worksheet",
    tags={"excel", "read", "chart"},
)
async def list_charts(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
) -> dict[str, Any]:
    """List all charts in an Excel worksheet."""
    try:
        backend = get_backend(file_path)

        ws = backend.get_sheet(sheet_name)

        charts = []
        for idx, chart in enumerate(ws._charts):
            title_text = _extract_chart_title(chart)
            charts.append({
                "index": idx,
                "title": title_text,
                "type": type(chart).__name__,
                "anchor": str(chart.anchor) if chart.anchor else None,
            })

        return {
            "success": True,
            "sheet_name": sheet_name,
            "charts": charts,
            "count": len(charts),
        }
    except Exception as e:
        logger.error("Error listing charts: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="delete_chart",
    description="Delete a chart from an Excel worksheet",
    tags={"excel", "write", "chart"},
)
async def delete_chart(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
    chart_index: Annotated[int, Field(description="Index of chart to delete (0-based)")],
) -> dict[str, Any]:
    """Delete a chart from an Excel worksheet."""
    try:
        backend = get_backend(file_path)

        ws = backend.get_sheet(sheet_name)

        if chart_index < 0 or chart_index >= len(ws._charts):
            return {
                "success": False,
                "error": f"Chart index {chart_index} out of range (0-{len(ws._charts) - 1})",
            }

        chart_title = _extract_chart_title(ws._charts[chart_index])
        del ws._charts[chart_index]

        backend.save()

        return {
            "success": True,
            "sheet_name": sheet_name,
            "deleted_chart_index": chart_index,
            "chart_title": chart_title,
        }
    except Exception as e:
        logger.error("Error deleting chart: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="modify_chart",
    description="Modify chart properties in an Excel worksheet",
    tags={"excel", "write", "chart"},
)
async def modify_chart(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    sheet_name: Annotated[str, Field(description="Worksheet name")],
    chart_index: Annotated[int, Field(description="Index of chart to modify (0-based)")],
    title: Annotated[str | None, Field(description="New chart title")] = None,
    style: Annotated[int | None, Field(description="Chart style (1-48)")] = None,
) -> dict[str, Any]:
    """Modify chart properties in an Excel worksheet."""
    try:
        backend = get_backend(file_path)

        ws = backend.get_sheet(sheet_name)

        if chart_index < 0 or chart_index >= len(ws._charts):
            return {
                "success": False,
                "error": f"Chart index {chart_index} out of range (0-{len(ws._charts) - 1})",
            }

        chart = ws._charts[chart_index]
        modified = []

        if title is not None:
            chart.title = title
            modified.append("title")

        if style is not None:
            if 1 <= style <= 48:
                chart.style = style
                modified.append("style")
            else:
                return {"success": False, "error": "Style must be between 1 and 48"}

        backend.save()

        return {
            "success": True,
            "sheet_name": sheet_name,
            "chart_index": chart_index,
            "modified_properties": modified,
        }
    except Exception as e:
        logger.error("Error modifying chart: %s", e)
        return {"success": False, "error": str(e)}


def _col_letter_to_index(cell_ref: str) -> int:
    """Convert column letter to 1-based index."""
    from openpyxl.utils import column_index_from_string

    col_letter = "".join(c for c in cell_ref if c.isalpha())
    return int(column_index_from_string(col_letter))


def _extract_chart_title(chart: Any) -> str:
    """Extract title text from chart object."""
    try:
        has_tx = hasattr(chart.title, "tx") and chart.title.tx
        has_rich = has_tx and hasattr(chart.title.tx, "rich") and chart.title.tx.rich
        if has_rich:
            paragraphs = chart.title.tx.rich.paragraphs
            if paragraphs and hasattr(paragraphs[0], "r"):
                runs = paragraphs[0].r
                if runs:
                    return str(runs[0].t)
        return str(chart.title) if chart.title else "Untitled"
    except Exception:
        return "Untitled"
