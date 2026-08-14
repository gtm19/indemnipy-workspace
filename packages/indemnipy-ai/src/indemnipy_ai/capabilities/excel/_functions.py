from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, cast

import polars as pl
from dateutil.parser import parse as parse_datetime
from oletools.olevba import VBA_Parser  # pyright: ignore[reportMissingTypeStubs]
from openpyxl.cell.cell import (
    Cell,
    MergedCell,
)
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import (
    Worksheet,
)

from ._parser_contract import (
    _ParsedWorkbookTable,
    _VbaAnalysisResult,
    _VbaMacro,
    _VbaParseResult,
)


def _oletools_vba_parser(filepath: Path) -> _VbaParseResult:
    parser = VBA_Parser(str(filepath))

    analysis_results: list[_VbaAnalysisResult] = []
    if parser.detect_vba_macros():
        analysis_results_raw = parser.analyze_macros()
        if analysis_results_raw is not None:
            for kw_type, keyword, description in analysis_results_raw:
                analysis_results.append(
                    _VbaAnalysisResult(
                        kw_type=kw_type,
                        keyword=keyword,
                        description=description,
                    )
                )

    macros: list[_VbaMacro] = []
    for filename, stream_path, vba_filename, vba_code in parser.extract_macros():
        macros.append(
            _VbaMacro(
                filename=cast(str, filename),
                stream_path=cast(str, stream_path),
                vba_filename=cast(str, vba_filename),
                vba_code=cast(str, vba_code),
            )
        )

    return _VbaParseResult(
        analysis_results=tuple(analysis_results),
        macros=tuple(macros),
    )


def _dataframe_from_table(worksheet: Worksheet, table_name: str) -> pl.DataFrame:
    """Convert an openpyxl Table to a Polars DataFrame."""
    # Get the table from the worksheet
    table: Table = worksheet.tables[table_name]
    # Get the range of the table
    range = table.ref

    return _dataframe_from_range(worksheet, range)


_COERCE_TYPES = {
    float: (float, int),
    int: (int, float),
}


@dataclass(frozen=True)
class DateParsingOptions:
    """Options for parsing dates in Excel data.

    Attributes:
        parse_dates: When ``True``, strings that can be unambiguously parsed as
            dates are converted to ``datetime`` objects.  Only applies to
            columns that already contain at least one date value.
        relaxed_about_day: When ``True``, accepts strings where the day is
            missing or ambiguous, substituting the 1st of the month.  When
            ``False`` the day must be present and unambiguous.
    """

    parse_dates: bool = True
    relaxed_about_day: bool = False


def _cleanse_data(
    data: dict[str, list[Any]], date_parsing_options: DateParsingOptions | None = None
) -> dict[str, list[Any]]:
    """Cleanse the data dictionary to ensure consistent types for each column."""
    date_parsing_options = date_parsing_options or DateParsingOptions()
    cleansed = {}

    for colname in data:
        values = data[colname]

        nonempty_indices = [i for i, v in enumerate(values) if v is not None]
        if not nonempty_indices:
            cleansed[colname] = values
            continue

        # get target type: being most common type in the column, ignoring None values
        type_counts: defaultdict[type, int] = defaultdict(int)
        for i in nonempty_indices:
            type_counts[type(values[i])] += 1
        target_type = max(type_counts, key=lambda t: type_counts[t])
        source_type = _COERCE_TYPES.get(target_type, (target_type,))

        if all(isinstance(values[i], source_type) for i in nonempty_indices):
            cleansed[colname] = values
        elif date_parsing_options.parse_dates and any(
            isinstance(values[i], (date, datetime)) for i in nonempty_indices
        ):
            non_date_indices = [
                i
                for i in nonempty_indices
                if not isinstance(values[i], (date, datetime))
            ]
            if non_date_indices:
                # if there are non-date values, we will attempt to parse them as dates
                cleansed[colname] = [
                    _gently_parse_datetime(
                        values[i],
                        relaxed_about_day=date_parsing_options.relaxed_about_day,
                    )
                    if i in non_date_indices
                    else values[i]
                    for i in range(len(values))
                ]
        else:
            cleansed[colname] = [str(v) if v is not None else None for v in values]

    return cleansed


def _dataframe_from_range(
    worksheet: Worksheet,
    range: str,
    date_parsing_options: DateParsingOptions | None = None,
) -> pl.DataFrame:
    """Convert an openpyxl range to a Polars DataFrame."""
    # Get the data from the range
    data = defaultdict(list)

    data_range: tuple[tuple[Cell | MergedCell, ...], ...] = worksheet[range]  # pyright: ignore[reportAssignmentType]

    for row in data_range[1:]:
        for colname, cell in zip(data_range[0], row):
            if isinstance(colname, MergedCell):
                raise ValueError(
                    f"Column name cell {colname.coordinate} is a merged cell, which is not supported."
                )
            if isinstance(cell, MergedCell):
                raise ValueError(
                    f"Data cell {cell.coordinate} is a merged cell, which is not supported."
                )
            if colname.value is not None:
                header = (
                    colname.value
                    if isinstance(colname.value, str)
                    else str(colname.value)
                )
                normalized = (
                    header.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
                )
                data[normalized].append(cell.value)

    df = pl.DataFrame(_cleanse_data(data, date_parsing_options=date_parsing_options))
    return df


def _openpyxl_table_parser(
    workbook: Workbook, date_parsing_options: DateParsingOptions | None = None
) -> dict[str, _ParsedWorkbookTable]:
    """Parse an Excel workbook and extract tables as Polars DataFrames.

    Args:
        workbook: An openpyxl Workbook instance.

    Returns:
        A dictionary mapping table names to _ParsedWorkbookTable instances.
    """

    tables: dict[str, _ParsedWorkbookTable] = {}

    for sheet_name in workbook.sheetnames:
        worksheet: Worksheet = workbook[sheet_name]
        for table_name in worksheet.tables:
            table: Table = worksheet.tables[table_name]
            range = table.ref
            dataframe = _dataframe_from_range(
                worksheet, range, date_parsing_options=date_parsing_options
            )
            tables[table_name] = _ParsedWorkbookTable(
                name=table_name,
                sheet_name=sheet_name,
                range=range,
                dataframe=dataframe,
            )

    return tables


def _gently_parse_datetime(value: Any, relaxed_about_day: bool = False) -> Any:  # pyright: ignore[reportAny, reportExplicitAny]
    """
    Safely parse a value to a datetime, returning the original value if parsing fails.

    If the input can be unambiguously parsed as a datetime, it will be returned as a datetime object. If not, the original value will be returned.

    If a date already (or parseable as such), a datetime object (at midnight) will be returned. If a datetime already,
    it will be returned as is. If a string, it will be parsed as a datetime if possible.

    Args:
        value: The value to parse.
        relaxed_about_day: If True, allows the day to be missing or invalid: in this case,
            the first of the month is used. If False, the day must be valid.

    Returns:
        The parsed datetime object, or the original value if parsing fails.

    """
    if isinstance(value, str):
        default_a = datetime(2000, 1, 1).replace(
            hour=0, minute=0, second=0, microsecond=0
        )
        default_b = default_a.replace(day=2)
        default_c = default_a.replace(month=2)

        try:
            a_day_first = parse_datetime(
                value, dayfirst=True, fuzzy=True, default=default_a
            )
            b_day_first = parse_datetime(
                value, dayfirst=True, fuzzy=True, default=default_b
            )
            a_month_first = parse_datetime(
                value, dayfirst=False, fuzzy=True, default=default_a
            )
            c_day_first = parse_datetime(
                value, dayfirst=True, fuzzy=True, default=default_c
            )

        except (ValueError, OverflowError):
            return value

        # existing ambiguity check: day/month order must agree either way
        if a_day_first != a_month_first:
            return value

        # month must be explicitly present in the string (e.g. "Back in 2021" has no month)
        month_present = a_day_first.month == c_day_first.month
        if not month_present:
            return value

        day_present = a_day_first.day == b_day_first.day
        if not day_present and not relaxed_about_day:
            return value  # day missing and we're not relaxed about it

        return a_day_first
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    return value  # pyright: ignore[reportAny]


def _print_df(df: pl.DataFrame, hide_shape: bool = False) -> str:
    """
    Print a Polars DataFrame with no formatting and all rows.

    Args:
        df: The Polars DataFrame to print.
        hide_shape: If True, the shape of the DataFrame will not be printed. Defaults to False.

    Returns:
        A string representation of the DataFrame.
    """
    with pl.Config(
        tbl_formatting="NOTHING", tbl_rows=-1, tbl_hide_dataframe_shape=hide_shape
    ):
        return str(df)
