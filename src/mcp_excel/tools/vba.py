"""MCP tools for Excel VBA macros."""

import logging
from typing import Annotated, Any

from fastmcp import FastMCP
from pydantic import Field

from ..backends.factory import create_backend
from ..utils.cache import shared_cache

logger = logging.getLogger(__name__)

# Create tools router
tools = FastMCP("Excel VBA Tools", mask_error_details=True)

# VBA Module types
MODULE_TYPES = ["standard", "class", "document"]

# VBA Code templates
VBA_TEMPLATES = {
    "hello_world": {
        "description": "Simple Hello World macro",
        'code': 'Sub HelloWorld()\n    MsgBox "Hello, World!"\nEnd Sub',
    },
    "loop_range": {
        "description": "Loop through a range and process cells",
        "code": """Sub LoopThroughRange()
    Dim cell As Range
    For Each cell In Range("A1:A10")
        If cell.Value <> "" Then
            cell.Interior.Color = vbYellow
        End If
    Next cell
End Sub""",
    },
    "filter_data": {
        "description": "AutoFilter example",
        "code": """Sub FilterData()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Range("A1").CurrentRegion.AutoFilter Field:=1, Criteria1:=">100"
End Sub""",
    },
    "create_chart": {
        "description": "Create a chart from data",
        "code": """Sub CreateChart()
    Dim chart As Chart
    Set chart = Charts.Add
    chart.SetSourceData Source:=ActiveSheet.UsedRange
    chart.ChartType = xlColumnClustered
End Sub""",
    },
    "copy_to_sheet": {
        "description": "Copy data to another sheet",
        "code": """Sub CopyToSheet()
    Dim src As Worksheet, dst As Worksheet
    Set src = Sheets("Source")
    Set dst = Sheets("Destination")
    src.UsedRange.Copy dst.Range("A1")
End Sub""",
    },
    "send_email": {
        "description": "Send email via Outlook",
        "code": """Sub SendEmail()
    Dim olApp As Object
    Dim olMail As Object
    Set olApp = CreateObject("Outlook.Application")
    Set olMail = olApp.CreateItem(0)
    With olMail
        .To = "recipient@example.com"
        .Subject = "Report"
        .Body = "Please see attached report."
        .Send
    End With
End Sub""",
    },
    "format_cells": {
        "description": "Format cells with colors and styles",
        "code": """Sub FormatCells()
    Dim rng As Range
    Set rng = Range("A1:D10")
    ' Header formatting
    With rng.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = vbWhite
    End With
    ' Data formatting
    With rng.Rows("2:" & rng.Rows.Count)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
End Sub""",
    },
    "pivot_table": {
        "description": "Create a PivotTable",
        "code": """Sub CreatePivotTable()
    Dim wsData As Worksheet
    Dim wsPivot As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Set wsData = Sheets("Data")
    Set wsPivot = Sheets.Add
    wsPivot.Name = "PivotTable"
    Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=wsData.UsedRange)
    Set pvt = pvtCache.CreatePivotTable( _
        TableDestination:=wsPivot.Range("A3"), _
        TableName:="PivotTable1")
End Sub""",
    },
    "validate_input": {
        "description": "Data validation with dropdown list",
        "code": """Sub AddDataValidation()
    Dim rng As Range
    Set rng = Range("A1:A100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="Option1,Option2,Option3"
        .ErrorTitle = "Invalid Input"
        .ErrorMessage = "Please select from the list."
    End With
End Sub""",
    },
    "workbook_events": {
        "description": "ThisWorkbook event handlers",
        "code": """' Place in ThisWorkbook module
Private Sub Workbook_Open()
    MsgBox "Welcome to " & ThisWorkbook.Name
End Sub

Private Sub Workbook_BeforeSave(Cancel As Boolean)
    If MsgBox("Save changes?", vbYesNo) = vbNo Then
        Cancel = True
    End If
End Sub""",
    },
    "error_handling": {
        "description": "Error handling template",
        "code": """Sub SafeMacro()
    On Error GoTo ErrorHandler
    ' Your code here
    Range("A1").Value = "Success"
    Exit Sub
ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub""",
    },
    "file_operations": {
        "description": "Read/write text files",
        "code": """Sub ExportToText()
    Dim fso As Object
    Dim ts As Object
    Dim rng As Range
    Dim cell As Range
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile("C:\\export.txt", True)
    Set rng = Range("A1:A" & Cells(Rows.Count, 1).End(xlUp).Row)
    For Each cell In rng
        ts.WriteLine cell.Value
    Next cell
    ts.Close
End Sub""",
    },
}


@tools.tool(
    name="list_vba_modules",
    description="List all VBA modules in an Excel workbook (.xlsm)",
    tags={"excel", "vba", "read"},
)
async def list_vba_modules(
    file_path: Annotated[str, Field(description="Absolute path to Excel file (.xlsm)")],
) -> dict[str, Any]:
    """List all VBA modules in an Excel workbook."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)
            shared_cache.put(file_path, backend)

        if not backend.has_macros():
            return {"success": False, "error": "File has no VBA macros"}

        modules = backend.list_vba_modules()
        return {"success": True, "modules": modules, "count": len(modules)}
    except Exception as e:
        logger.error("Error listing VBA modules: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="get_vba_code",
    description="Get VBA source code from a module",
    tags={"excel", "vba", "read"},
)
async def get_vba_code(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    module_name: Annotated[str, Field(description="VBA module name")],
) -> dict[str, Any]:
    """Get VBA source code from a module."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)
            shared_cache.put(file_path, backend)

        code = backend.get_vba_code(module_name)
        return {"success": True, "module_name": module_name, "code": code}
    except Exception as e:
        logger.error("Error getting VBA code: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="set_vba_code",
    description="Set/replace VBA source code in a module",
    tags={"excel", "vba", "write"},
)
async def set_vba_code(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    module_name: Annotated[str, Field(description="VBA module name")],
    code: Annotated[str, Field(description="VBA source code")],
) -> dict[str, Any]:
    """Set/replace VBA source code in a module."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)

        backend.set_vba_code(module_name, code)
        backend.save()
        shared_cache.put(file_path, backend)

        return {"success": True, "module_name": module_name, "code_length": len(code)}
    except Exception as e:
        logger.error("Error setting VBA code: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="add_vba_module",
    description="Add a new VBA module to the workbook",
    tags={"excel", "vba", "write"},
)
async def add_vba_module(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    module_name: Annotated[str, Field(description="Name for the new module")],
    code: Annotated[str, Field(description="VBA source code")] = "",
    module_type: Annotated[
        str, Field(description="Module type: standard, class, document")
    ] = "standard",
) -> dict[str, Any]:
    """Add a new VBA module to the workbook."""
    try:
        if module_type not in MODULE_TYPES:
            return {
                "success": False,
                "error": f"Invalid module type: {module_type}. Use: {MODULE_TYPES}",
            }

        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)

        backend.add_vba_module(module_name, code, module_type)
        backend.save()
        shared_cache.put(file_path, backend)

        return {"success": True, "module_name": module_name, "module_type": module_type}
    except Exception as e:
        logger.error("Error adding VBA module: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="delete_vba_module",
    description="Delete a VBA module from the workbook",
    tags={"excel", "vba", "write"},
)
async def delete_vba_module(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    module_name: Annotated[str, Field(description="VBA module name to delete")],
) -> dict[str, Any]:
    """Delete a VBA module from the workbook."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)

        backend.delete_vba_module(module_name)
        backend.save()
        shared_cache.put(file_path, backend)

        return {"success": True, "deleted_module": module_name}
    except Exception as e:
        logger.error("Error deleting VBA module: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="rename_vba_module",
    description="Rename a VBA module",
    tags={"excel", "vba", "write"},
)
async def rename_vba_module(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    old_name: Annotated[str, Field(description="Current module name")],
    new_name: Annotated[str, Field(description="New module name")],
) -> dict[str, Any]:
    """Rename a VBA module."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)

        backend.rename_vba_module(old_name, new_name)
        backend.save()
        shared_cache.put(file_path, backend)

        return {"success": True, "old_name": old_name, "new_name": new_name}
    except Exception as e:
        logger.error("Error renaming VBA module: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="run_macro",
    description="Execute a VBA macro by name",
    tags={"excel", "vba", "execute"},
)
async def run_macro(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    macro_name: Annotated[
        str, Field(description="Macro name (e.g., 'Module1.MyMacro' or 'MyMacro')")
    ],
    args: Annotated[list[Any] | None, Field(description="Arguments to pass to the macro")] = None,
) -> dict[str, Any]:
    """Execute a VBA macro by name."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)
            shared_cache.put(file_path, backend)

        macro_args = args if args is not None else []
        result = backend.run_macro(macro_name, *macro_args)
        return {"success": True, "macro_name": macro_name, "result": result}
    except Exception as e:
        logger.error("Error running macro: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="list_macros",
    description="List all Sub/Function procedures in VBA project",
    tags={"excel", "vba", "read"},
)
async def list_macros(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
) -> dict[str, Any]:
    """List all Sub/Function procedures in VBA project."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)
            shared_cache.put(file_path, backend)

        macros = backend.list_macros()
        return {"success": True, "macros": macros, "count": len(macros)}
    except Exception as e:
        logger.error("Error listing macros: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="get_vba_templates",
    description="Get VBA code templates for common tasks",
    tags={"excel", "vba", "template"},
)
async def get_vba_templates(
    template_type: Annotated[
        str | None, Field(description="Template type (or None for all)")
    ] = None,
) -> dict[str, Any]:
    """Get VBA code templates for common tasks."""
    try:
        if template_type:
            if template_type not in VBA_TEMPLATES:
                return {"success": False, "error": f"Template not found: {template_type}"}
            return {
                "success": True,
                "templates": {template_type: VBA_TEMPLATES[template_type]},
            }
        return {"success": True, "templates": VBA_TEMPLATES}
    except Exception as e:
        return {"success": False, "error": str(e)}


@tools.tool(
    name="validate_vba_code",
    description="Validate VBA code syntax (basic validation)",
    tags={"excel", "vba", "validate"},
)
async def validate_vba_code(
    code: Annotated[str, Field(description="VBA code to validate")],
) -> dict[str, Any]:
    """Validate VBA code syntax (basic validation)."""
    errors: list[str] = []
    warnings: list[str] = []
    subs: list[str] = []
    functions: list[str] = []

    lines = code.split("\n")

    for i, line in enumerate(lines, 1):
        stripped = line.strip()

        # Skip comments and empty lines
        if stripped.startswith("'") or not stripped:
            continue

        # Check for Sub/Function declarations
        if stripped.startswith("Sub ") or stripped.startswith("Public Sub "):
            name = stripped.split("Sub ")[1].split("(")[0].strip()
            subs.append(name)
        elif stripped.startswith("Function ") or stripped.startswith("Public Function "):
            name = stripped.split("Function ")[1].split("(")[0].strip()
            functions.append(name)

        # Check for variable declaration without type
        if "Dim " in stripped and " As " not in stripped and not stripped.startswith("'"):
            warnings.append(f"Line {i}: Variable declaration missing type")

    # Check for matching Sub/End Sub
    if subs:
        end_subs = sum(1 for line in lines if line.strip() == "End Sub")
        if end_subs != len(subs):
            errors.append(f"Mismatched Sub/End Sub: {len(subs)} Sub vs {end_subs} End Sub")

    # Check for matching Function/End Function
    if functions:
        end_functions = sum(1 for line in lines if line.strip() == "End Function")
        if end_functions != len(functions):
            errors.append(
                f"Mismatched Function/End Function: {len(functions)} Function vs "
                f"{end_functions} End Function"
            )

    return {
        "success": True,
        "valid": len(errors) == 0,
        "errors": errors,
        "warnings": warnings,
        "subs": subs,
        "functions": functions,
    }


@tools.tool(
    name="import_vba_module",
    description="Import VBA code from a string to a module",
    tags={"excel", "vba", "write"},
)
async def import_vba_module(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    module_name: Annotated[str, Field(description="Target module name")],
    code: Annotated[str, Field(description="VBA code to import")],
    module_type: Annotated[
        str, Field(description="Module type: standard, class, document")
    ] = "standard",
) -> dict[str, Any]:
    """Import VBA code from a string to a module."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)

        # Check if module exists
        try:
            backend.get_vba_code(module_name)
            # Module exists, update it
            backend.set_vba_code(module_name, code)
        except ValueError:
            # Module doesn't exist, create it
            backend.add_vba_module(module_name, code, module_type)

        backend.save()
        shared_cache.put(file_path, backend)

        return {
            "success": True,
            "module_name": module_name,
            "code_length": len(code),
        }
    except Exception as e:
        logger.error("Error importing VBA module: %s", e)
        return {"success": False, "error": str(e)}


@tools.tool(
    name="export_vba_module",
    description="Export VBA code from a module to a string",
    tags={"excel", "vba", "read"},
)
async def export_vba_module(
    file_path: Annotated[str, Field(description="Absolute path to Excel file")],
    module_name: Annotated[str, Field(description="VBA module name to export")],
) -> dict[str, Any]:
    """Export VBA code from a module to a string."""
    try:
        backend = shared_cache.get(file_path)
        if backend is None:
            backend = create_backend(file_path)
            backend.open(file_path)
            shared_cache.put(file_path, backend)

        code = backend.get_vba_code(module_name)
        return {
            "success": True,
            "module_name": module_name,
            "code": code,
            "code_length": len(code),
        }
    except Exception as e:
        logger.error("Error exporting VBA module: %s", e)
        return {"success": False, "error": str(e)}
