"""Small helpers shared by the generic workbook transformation steps."""

from copy import copy
from datetime import date, datetime

import pandas as pd


def is_missing(value) -> bool:
    """Return True for a scalar pandas/Excel missing value."""
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def text_value(value) -> str:
    """Return a trimmed string without exposing pandas' missing-value text."""
    return "" if is_missing(value) else str(value).strip()


def format_date(value):
    """Return an Excel-friendly date value or a blank string."""
    if is_missing(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, (date, datetime)):
        return value
    return value


def mapping_value(mapping_row, column_name, default=""):
    """Read a named mapping value, returning a default when it is unavailable."""
    if mapping_row is None or column_name not in mapping_row:
        return default
    value = mapping_row[column_name]
    return default if is_missing(value) else value


def normalise_boolean(value):
    """Return common Boolean representations as True/False text."""
    if is_missing(value):
        return ""
    if isinstance(value, bool) or type(value).__name__ == "bool_":
        return str(bool(value))
    normalised = str(value).strip().casefold()
    if normalised in {"true", "1", "yes"}:
        return "True"
    if normalised in {"false", "0", "no"}:
        return "False"
    return value


def template_row_styles(worksheet, header_row=1):
    """Snapshot the first data-row styles before template rows are cleared."""
    style_row = min(header_row + 1, worksheet.max_row)
    return {
        column: copy(worksheet.cell(style_row, column)._style)
        for column in range(1, worksheet.max_column + 1)
        if worksheet.cell(style_row, column).has_style
    }


def clear_data_rows(worksheet, header_row=1):
    """Remove existing data rows while retaining the worksheet header."""
    first_data_row = header_row + 1
    if worksheet.max_row >= first_data_row:
        worksheet.delete_rows(first_data_row, worksheet.max_row - header_row)


def apply_template_styles(worksheet, row_number, styles):
    """Apply a saved template style to a newly created data row."""
    for column, style in styles.items():
        worksheet.cell(row_number, column)._style = copy(style)
