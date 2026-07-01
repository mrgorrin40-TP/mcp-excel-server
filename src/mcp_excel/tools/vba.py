"""MCP tools for Excel VBA macros."""

import logging
from typing import Annotated, Any

from fastmcp import FastMCP
from pydantic import Field

from ..utils.backend import get_backend
from ..utils.vba_templates import VBA_TEMPLATES
from ..utils.vba_validator import validate_vba_code as validate_code

logger = logging.getLogger(__name__)

# Create tools router
tools = FastMCP("Excel VBA Tools", mask_error_details=True)

# VBA Module types
MODULE_TYPES = ["standard", "class", "document"]


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
        backend = get_backend(file_path)

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
        backend = get_backend(file_path)

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
        backend = get_backend(file_path)

        backend.set_vba_code(module_name, code)
        backend.save()

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

        backend = get_backend(file_path)

        backend.add_vba_module(module_name, code, module_type)
        backend.save()

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
        backend = get_backend(file_path)

        backend.delete_vba_module(module_name)
        backend.save()

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
        backend = get_backend(file_path)

        backend.rename_vba_module(old_name, new_name)
        backend.save()

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
        backend = get_backend(file_path)

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
        backend = get_backend(file_path)

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
    result = validate_code(code)
    return {
        "success": True,
        "valid": result.valid,
        "errors": result.errors,
        "warnings": result.warnings,
        "subs": result.subs,
        "functions": result.functions,
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
        backend = get_backend(file_path)

        # Check if module exists
        try:
            backend.get_vba_code(module_name)
            # Module exists, update it
            backend.set_vba_code(module_name, code)
        except ValueError:
            # Module doesn't exist, create it
            backend.add_vba_module(module_name, code, module_type)

        backend.save()

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
        backend = get_backend(file_path)

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
