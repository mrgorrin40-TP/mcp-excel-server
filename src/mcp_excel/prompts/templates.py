"""MCP prompts for Excel analysis workflows."""

from typing import Any

ANALYSIS_PROMPTS = {
    "explore_workbook": {
        "description": "Explore and summarize an Excel workbook",
        "template": (
            "Analyze the Excel file at {file_path}.\n\n"
            "Steps:\n"
            "1. Use describe_workbook to get an overview\n"
            "2. Use list_sheets to see all worksheets\n"
            "3. For each sheet, use get_sheet_info to understand structure\n"
            "4. Use read_range to sample data from key sheets\n"
            "5. Provide a summary of the workbook contents"
        ),
    },
    "analyze_data": {
        "description": "Perform statistical analysis on worksheet data",
        "template": (
            "Analyze the data in sheet '{sheet_name}' of {file_path}.\n\n"
            "Steps:\n"
            "1. Use get_sheet_info to understand columns and data types\n"
            "2. For numeric columns, use get_column_stats to get statistics\n"
            "3. Use group_by to identify patterns and groupings\n"
            "4. Use filter_rows to identify outliers or specific conditions\n"
            "5. Summarize key findings and insights"
        ),
    },
    "compare_sheets": {
        "description": "Compare data between two worksheets",
        "template": (
            "Compare the data in sheets '{sheet1}' and '{sheet2}' of {file_path}.\n\n"
            "Steps:\n"
            "1. Use get_sheet_info on both sheets to understand structure\n"
            "2. Use read_range to sample data from both sheets\n"
            "3. Identify common columns for comparison\n"
            "4. Use get_column_stats on matching columns\n"
            "5. Use group_by to compare aggregated data\n"
            "6. Summarize similarities and differences"
        ),
    },
    "create_report": {
        "description": "Create a summary report from raw data",
        "template": (
            "Create a summary report from the data in {file_path}.\n\n"
            "Steps:\n"
            "1. Use describe_workbook to understand the file structure\n"
            "2. Use get_sheet_info on the main data sheet\n"
            "3. Use get_column_stats for key numeric columns\n"
            "4. Use group_by to aggregate data by categories\n"
            "5. Use create_sheet to add a new 'Summary' sheet\n"
            "6. Use write_cells to add summary headers and data\n"
            "7. Use write_formula to add calculations\n"
            "8. Save the file with the new summary sheet"
        ),
    },
    "data_quality_check": {
        "description": "Check data quality and identify issues",
        "template": (
            "Perform a data quality check on sheet '{sheet_name}' of {file_path}.\n\n"
            "Steps:\n"
            "1. Use get_sheet_info to understand the data structure\n"
            "2. For each column, use get_column_stats to check for nulls\n"
            "3. Use filter_rows with 'is_null' operator to find missing values\n"
            "4. Use search_cells to find potential data issues\n"
            "5. Identify duplicates, outliers, or inconsistencies\n"
            "6. Provide a data quality report with recommendations"
        ),
    },
}


def get_prompt(name: str, **kwargs: Any) -> dict[str, Any]:
    """Get a formatted prompt template.

    Args:
        name: Prompt name
        **kwargs: Template variables

    Returns:
        Formatted prompt dictionary
    """
    if name not in ANALYSIS_PROMPTS:
        return {"error": f"Prompt '{name}' not found"}

    prompt = ANALYSIS_PROMPTS[name]

    try:
        formatted_text = prompt["template"].format(**kwargs)
    except KeyError as e:
        return {"error": f"Missing required variable: {e}"}

    return {
        "name": name,
        "description": prompt["description"],
        "text": formatted_text,
    }


def list_prompts() -> dict[str, Any]:
    """List all available prompts.

    Returns:
        Dictionary with available prompts
    """
    prompts = []
    for name, prompt in ANALYSIS_PROMPTS.items():
        prompts.append({
            "name": name,
            "description": prompt["description"],
        })

    return {
        "prompts": prompts,
        "count": len(prompts),
    }
