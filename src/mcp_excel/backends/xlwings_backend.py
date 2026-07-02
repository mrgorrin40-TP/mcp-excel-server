"""Xlwings backend for Excel VBA macro operations."""

import logging
from pathlib import Path
from typing import Any

from .base import ExcelBackend

logger = logging.getLogger(__name__)


class XlwingsBackend(ExcelBackend):
    """Excel backend using xlwings for VBA macro support.

    Requires Excel to be installed on the system.
    Supports Windows and macOS.
    """

    def __init__(self) -> None:
        self._app: Any = None
        self._book: Any = None
        self._file_path: Path | None = None

    def open(self, file_path: str | Path) -> None:
        """Open an Excel file with xlwings."""
        try:
            import xlwings as xw
        except ImportError as err:
            raise ImportError(
                "xlwings is required for VBA support. "
                "Install with: pip install 'mcp-excel-server[vba]'"
            ) from err

        self._file_path = Path(file_path)

        if not self._file_path.exists():
            raise FileNotFoundError(f"File not found: {self._file_path}")

        suffix = self._file_path.suffix.lower()
        if suffix not in (".xlsm", ".xlsb", ".xlam", ".xlsx"):
            raise ValueError(f"Unsupported file format: {suffix}")

        logger.info("Opening workbook with xlwings: %s", self._file_path)
        self._app = xw.App(visible=False)
        self._book = self._app.books.open(str(self._file_path))

    def close(self) -> None:
        """Close workbook and Excel application."""
        if self._book:
            logger.info("Closing workbook")
            self._book.close()
            self._book = None
        if self._app:
            self._app.quit()
            self._app = None
        self._file_path = None

    def save(self, file_path: str | Path | None = None) -> None:
        """Save the workbook."""
        if not self._book:
            raise ValueError("No workbook open")

        save_path = Path(file_path) if file_path else self._file_path
        if not save_path:
            raise ValueError("No file path specified")

        logger.info("Saving workbook to: %s", save_path)
        self._book.save(str(save_path))
        self._file_path = save_path

    def get_sheet_names(self) -> list[str]:
        """Get list of all sheet names."""
        if not self._book:
            raise ValueError("No workbook open")

        return [sheet.name for sheet in self._book.sheets]

    def get_sheet(self, sheet_name: str) -> Any:
        """Get a worksheet by name."""
        if not self._book:
            raise ValueError("No workbook open")

        for sheet in self._book.sheets:
            if sheet.name == sheet_name:
                return sheet

        raise ValueError(f"Sheet not found: {sheet_name}")

    def create_sheet(self, sheet_name: str) -> Any:
        """Create a new worksheet."""
        if not self._book:
            raise ValueError("No workbook open")

        for sheet in self._book.sheets:
            if sheet.name == sheet_name:
                raise ValueError(f"Sheet already exists: {sheet_name}")

        logger.info("Creating sheet: %s", sheet_name)
        return self._book.sheets.add(name=sheet_name)

    def delete_sheet(self, sheet_name: str) -> None:
        """Delete a worksheet."""
        if not self._book:
            raise ValueError("No workbook open")

        sheet = self.get_sheet(sheet_name)
        if len(self._book.sheets) <= 1:
            raise ValueError("Cannot delete the last sheet")

        logger.info("Deleting sheet: %s", sheet_name)
        sheet.delete()

    def read_cell(self, sheet_name: str, cell: str) -> Any:
        """Read a single cell value."""
        sheet = self.get_sheet(sheet_name)
        return sheet.range(cell).value

    def write_cell(self, sheet_name: str, cell: str, value: Any) -> None:
        """Write a value to a cell."""
        sheet = self.get_sheet(sheet_name)
        logger.debug("Writing %s to %s!%s", value, sheet_name, cell)
        sheet.range(cell).value = value

    def read_range(self, sheet_name: str, range_str: str) -> list[list[Any]]:
        """Read a range of cells."""
        sheet = self.get_sheet(sheet_name)
        result: list[list[Any]] = sheet.range(range_str).value
        return result

    def write_range(self, sheet_name: str, range_str: str, values: list[list[Any]]) -> None:
        """Write values to a range."""
        sheet = self.get_sheet(sheet_name)
        logger.debug("Writing data to %s!%s", sheet_name, range_str)
        sheet.range(range_str).value = values

    def get_used_range(self, sheet_name: str) -> str:
        """Get the used range of a worksheet."""
        sheet = self.get_sheet(sheet_name)
        used_range = sheet.used_range
        if used_range:
            result: str = used_range.address
            return result
        return "A1"

    def get_cell_type(self, sheet_name: str, cell: str) -> str:
        """Get the data type of a cell."""
        sheet = self.get_sheet(sheet_name)
        value = sheet.range(cell).value

        if value is None:
            return "null"

        from datetime import date, datetime

        if isinstance(value, datetime | date):
            return "date"

        if isinstance(value, bool):
            return "boolean"

        if isinstance(value, int | float):
            return "number"

        return "string"

    def get_max_row(self, sheet_name: str) -> int:
        """Get the maximum row number with data in a worksheet."""
        sheet = self.get_sheet(sheet_name)
        used_range = sheet.used_range
        if used_range:
            last_cell = used_range.last_cell
            return int(last_cell.row)
        return 0

    def get_max_column(self, sheet_name: str) -> int:
        """Get the maximum column number with data in a worksheet."""
        sheet = self.get_sheet(sheet_name)
        used_range = sheet.used_range
        if used_range:
            last_cell = used_range.last_cell
            return int(last_cell.column)
        return 0

    # VBA-specific methods

    def has_macros(self) -> bool:
        """Check if workbook contains VBA macros."""
        if not self._book:
            raise ValueError("No workbook open")

        try:
            # Access VBProject to check for VBA
            _ = self._book.api.VBProject
            return True
        except Exception:
            return False

    def list_vba_modules(self) -> list[dict[str, Any]]:
        """List all VBA modules in the workbook."""
        if not self._book:
            raise ValueError("No workbook open")

        modules = []
        try:
            vb_project = self._book.api.VBProject
            for component in vb_project.VBComponents:
                modules.append({
                    "name": component.Name,
                    "type": self._get_component_type(component.Type),
                    "description": component.Description or "",
                })
        except Exception as e:
            logger.error("Error listing VBA modules: %s", e)
            raise ValueError(f"Cannot access VBA project: {e}") from e

        return modules

    def get_vba_code(self, module_name: str) -> str:
        """Get VBA source code from a module."""
        if not self._book:
            raise ValueError("No workbook open")

        try:
            vb_project = self._book.api.VBProject
            component = vb_project.VBComponents(module_name)
            code_module = component.CodeModule
            if code_module.CountOfLines > 0:
                result: str = code_module.Lines(1, code_module.CountOfLines)
                return result
            return ""
        except Exception as e:
            raise ValueError(f"Module '{module_name}' not found: {e}") from e

    def set_vba_code(self, module_name: str, code: str) -> None:
        """Set/replace VBA source code in a module."""
        if not self._book:
            raise ValueError("No workbook open")

        try:
            vb_project = self._book.api.VBProject
            component = vb_project.VBComponents(module_name)
            code_module = component.CodeModule

            # Clear existing code
            if code_module.CountOfLines > 0:
                code_module.DeleteLines(1, code_module.CountOfLines)

            # Add new code
            if code:
                code_module.AddFromString(code)
        except Exception as e:
            raise ValueError(f"Error setting code for '{module_name}': {e}") from e

    def add_vba_module(self, name: str, code: str = "", module_type: str = "standard") -> None:
        """Add a new VBA module."""
        if not self._book:
            raise ValueError("No workbook open")

        # vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2, vbext_ct_Document = 100
        type_map = {
            "standard": 1,
            "class": 2,
            "document": 100,
        }

        if module_type not in type_map:
            raise ValueError(f"Invalid module type: {module_type}. Use: {list(type_map.keys())}")

        try:
            vb_project = self._book.api.VBProject
            component = vb_project.VBComponents.Add(type_map[module_type])
            component.Name = name

            if code:
                component.CodeModule.AddFromString(code)
        except Exception as e:
            raise ValueError(f"Error adding module '{name}': {e}") from e

    def delete_vba_module(self, name: str) -> None:
        """Delete a VBA module."""
        if not self._book:
            raise ValueError("No workbook open")

        try:
            vb_project = self._book.api.VBProject
            component = vb_project.VBComponents(name)
            vb_project.VBComponents.Remove(component)
        except Exception as e:
            raise ValueError(f"Error deleting module '{name}': {e}") from e

    def rename_vba_module(self, old_name: str, new_name: str) -> None:
        """Rename a VBA module."""
        if not self._book:
            raise ValueError("No workbook open")

        try:
            vb_project = self._book.api.VBProject
            component = vb_project.VBComponents(old_name)
            component.Name = new_name
        except Exception as e:
            raise ValueError(f"Error renaming module '{old_name}': {e}") from e

    def run_macro(self, macro_name: str, *args: Any) -> Any:
        """Execute a VBA macro by name with timeout support."""
        if not self._book:
            raise ValueError("No workbook open")

        from concurrent.futures import ThreadPoolExecutor
        from concurrent.futures import TimeoutError as FuturesTimeout

        from ..config import settings

        def _execute() -> Any:
            macro = self._book.app.macro(macro_name)
            if args:
                return macro(*args)
            return macro()

        timeout = settings.vba_macro_timeout

        try:
            with ThreadPoolExecutor(max_workers=1) as executor:
                future = executor.submit(_execute)
                return future.result(timeout=timeout)
        except FuturesTimeout as err:
            raise ValueError(
                f"Macro '{macro_name}' timed out after {timeout} seconds"
            ) from err
        except Exception as e:
            if "timed out" in str(e).lower():
                raise
            raise ValueError(f"Error running macro '{macro_name}': {e}") from e

    def list_macros(self) -> list[dict[str, Any]]:
        """List all Sub/Function procedures in VBA project."""
        if not self._book:
            raise ValueError("No workbook open")

        macros = []
        try:
            vb_project = self._book.api.VBProject
            for component in vb_project.VBComponents:
                code_module = component.CodeModule
                if code_module.CountOfLines > 0:
                    code = code_module.Lines(1, code_module.CountOfLines)
                    for line in code.split("\n"):
                        line = line.strip()
                        if line.startswith("Public Sub ") or line.startswith("Private Sub "):
                            name = line.split("Sub ")[1].split("(")[0].strip()
                            macros.append({
                                "name": name,
                                "type": "Sub",
                                "module": component.Name,
                                "is_public": line.startswith("Public"),
                            })
                        elif line.startswith("Public Function ") or line.startswith(
                            "Private Function "
                        ):
                            name = line.split("Function ")[1].split("(")[0].strip()
                            macros.append({
                                "name": name,
                                "type": "Function",
                                "module": component.Name,
                                "is_public": line.startswith("Public"),
                            })
        except Exception as e:
            logger.error("Error listing macros: %s", e)

        return macros

    def _get_component_type(self, type_code: int) -> str:
        """Convert component type code to string."""
        types = {
            1: "standard",
            2: "class",
            3: "form",
            100: "document",
        }
        return types.get(type_code, "unknown")
