from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol

import polars as pl

from indemnipy_ai.capabilities.excel._parser_contract import (
    _VbaAnalysisResult,
    _VbaMacro,
    _VbaParser,
)


@dataclass
class VbaSummary:
    """Aggregates extracted VBA macros and heuristic analysis for one file.

    Produced by [`WorkbookProtocol.vba_summary`][indemnipy_ai.capabilities.excel.WorkbookProtocol.vba_summary]
    when macros are detected
    in a workbook.  Use [`to_md`][indemnipy_ai.capabilities.excel.VbaSummary.to_md] to render the summary as a Markdown
    document suitable for passing to an agent.

    Attributes:
        filepath: Path to the workbook that was inspected.
        analysis_results: Heuristic findings returned by macro analysis (keyword
            type, keyword, and description).
        macros: Extracted VBA macro streams (filename, stream path, VBA
            filename, and source code).
    """

    filepath: Path
    analysis_results: list[_VbaAnalysisResult]
    macros: list[_VbaMacro]

    @classmethod
    def from_file(cls, filepath: Path, parser: _VbaParser) -> "VbaSummary":
        """Build a summary from an Excel workbook or OLE document.

        Args:
            filepath: Path to the workbook to inspect.

        Returns:
            A summary containing any detected analysis findings and extracted
                VBA macro streams.
        """
        parse_result = parser(filepath)
        return cls(
            filepath=filepath,
            analysis_results=list(parse_result.analysis_results),
            macros=list(parse_result.macros),
        )

    def to_md(self) -> str:
        """Render the summary as Markdown.

        Returns:
            A Markdown document containing analysis results followed by each
                extracted macro and its VBA source code.
        """

        content: str = f"# VBA Macros in {self.filepath.name}\n"

        if self.analysis_results:
            content += f"\n## Analysis Results\nThere are {len(self.analysis_results)} observations.\n"
            for i, result in enumerate(self.analysis_results):
                content += f"{i + 1}. Type: {result.kw_type}, Keyword: {result.keyword}, Description: {result.description}\n"

        for i, macro in enumerate(self.macros):
            content += f"\n## Macro {i + 1}\n"
            content += f"Filename: {macro.filename}\n"
            content += f"Stream Path: {macro.stream_path}\n"
            content += f"VBA Filename: {macro.vba_filename}\n\n"
            content += f"""### VBA Code:
```vba
{macro.vba_code}
```
"""
        return content


@dataclass
class WorkbookTable:
    """A single table extracted from an Excel workbook, or produced by a query.

    Workbook tables come from named Excel tables or ranges registered with
    ``add_table_from_range``.  Derived tables are created by
    ``query_store_and_preview`` and live in
    [`ExcelRuntimeState.derived_tables`][indemnipy_ai.capabilities.excel.ExcelRuntimeState.derived_tables].
    For both types, the ``dataframe`` attribute;
    for those, ``sheet_name`` and ``range`` are empty strings.
    """

    name: str
    """The name of the table as defined in the workbook."""
    sheet_name: str
    """The name of the sheet where the table is located."""
    range: str
    """The range of cells occupied by the table, e.g. ``"A1:C10"``."""
    dataframe: pl.DataFrame = field(repr=False)
    """The table's data as a Polars DataFrame."""


@dataclass
class WorkbookSheet:
    """Metadata for a single worksheet in a loaded workbook.

    Returned via :attr:`WorkbookProtocol.sheets`.  The ``tables`` list
    contains only *named* Excel tables.  Sheets with tabular data that has not
    been formally defined as a table will still appear here but with an empty
    ``tables`` list; use ``get_range`` to read their raw contents.
    """

    name: str
    """The name of the sheet in the workbook."""
    range: str
    """The full cell range occupied by the sheet, e.g. ``"A1:Z100"``."""
    freeze_panes: str | None
    """Freeze panes anchor cell, e.g. ``"B2"``, or ``None``.  Its presence may
    indicate tabular data even without a formally defined table."""
    min_column: int
    """1-based index of the first used column."""
    min_row: int
    """1-based index of the first used row."""
    max_column: int
    """1-based index of the last used column."""
    max_row: int
    """1-based index of the last used row."""
    state: str
    """Visibility state of the sheet: ``'visible'`` or ``'hidden'``."""
    tables: list[WorkbookTable]
    """Named tables defined on this sheet."""


class WorkbookProtocol(Protocol):
    """Structural protocol for a loaded Excel workbook.

    Instances are created internally and stored in
    :attr:`ExcelRuntimeState.workbooks <indemnipy_ai.capabilities.excel.ExcelRuntimeState.workbooks>`.
    You will not normally construct these directly, but you can read them after
    a run to inspect what the agent loaded.
    """

    filepath: Path
    """Path to the source file."""

    @property
    def file_name(self) -> str:
        """File name without directory path.

        Returns:
            The name of the workbook file as a string.
        """
        ...

    @property
    def vba_summary(self) -> VbaSummary | None:
        """VBA macro summary, or ``None`` if the workbook contains no macros.

        Returns:
            A `VbaSummary` instance if macros are present, otherwise
                ``None``.
        """
        ...

    @property
    def sheets(self) -> list[WorkbookSheet]:
        """All worksheets in the workbook.

        Returns:
            A list of :class:`WorkbookSheet` instances.
        """
        ...

    def get_range(self, sheet_name: str, range_str: str) -> list[list[Any]]:
        """Return raw cell values for a given range.

        Args:
            sheet_name: Name of the sheet to read from.
            range_str: Excel-style range string, e.g. ``"A1:C3"``.

        Returns:
            A list of rows, each row being a list of cell values.
        """
        ...

    def add_table_from_range(
        self, sheet_name: str, range_str: str, table_name: str
    ) -> None:
        """Register a cell range as a named table on a sheet.

        Args:
            sheet_name: Name of the sheet containing the range.
            range_str: Excel-style range string, e.g. ``"A1:C3"``.
            table_name: Name to assign to the new table.  Must be unique within
                the sheet.
        """
        ...

    def agent_summary(self) -> str:
        """Return a short text summary of the workbook's sheets and tables.

        Returns:
            A string suitable for passing to an agent as context.
        """
        ...
