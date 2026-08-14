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
    """Aggregates extracted VBA macros and parser analysis for one file.

    Attributes:
        filepath: Workbook path that was inspected.
        analysis_results: Heuristic findings returned by macro analysis.
        macros: Extracted VBA macro streams.
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
    name: str
    """The name of the table as defined in the workbook."""
    sheet_name: str
    """The name of the sheet where the table is located."""
    range: str
    """The range of cells occupied by the table. e.g. "A1:C10"."""
    dataframe: pl.DataFrame = field(repr=False)
    """The table's data represented as a Polars DataFrame."""


@dataclass
class WorkbookSheet:
    name: str
    """The name of the sheet in the workbook."""
    range: str
    """The range of cells occupied by the sheet. e.g. "A1:C10"."""
    freeze_panes: str | None
    """The cell reference for the freeze panes position, if any. e.g. "A1". The presence of this might indicate tabular information (even if not formally defined as a table)."""
    min_column: int
    """The minimum column index of the table (1-based)."""
    min_row: int
    """The minimum row index of the table (1-based)."""
    max_column: int
    """The maximum column index of the table (1-based)."""
    max_row: int
    """The maximum row index of the table (1-based)."""
    state: str
    """The visibility state of the sheet. Can be 'visible', 'hidden'."""
    tables: list[WorkbookTable]
    """A list of tables defined in the sheet."""


class WorkbookProtocol(Protocol):
    filepath: Path

    @property
    def file_name(self) -> str:
        """
        Return the name of the workbook file.

        Returns:
            The name of the workbook file as a string.
        """
        ...

    @property
    def vba_summary(self) -> VbaSummary | None:
        """
        Return a summary of any detected VBA macros and analysis results.

        Returns:
            A VbaSummary instance if macros are present, otherwise None.
        """
        ...

    @property
    def sheets(self) -> list[WorkbookSheet]:
        """
        Return a list of sheet names in the workbook.

        Returns:
            A list of WorkbookSheet instances representing the sheets in the workbook.
        """
        ...

    def get_range(self, sheet_name: str, range_str: str) -> list[list[Any]]:
        """
        Get the values in a specified range of a sheet.

        Args:
            sheet_name: The name of the sheet.
            range_str: The Excel-style range string (e.g., "A1:C3").

        Returns:
            A list of lists containing the values in the specified range.
        """
        ...

    def add_table_from_range(
        self, sheet_name: str, range_str: str, table_name: str
    ) -> None:
        """
        Add a new table to the workbook, given a sheet name, range, and table name.

        This method should internally create a new WorkbookTable instance and add it to the workbook's tables dictionary.

        Args:
            sheet_name: The name of the sheet containing the new table.
            range_str: The Excel-style range string for the new table (e.g., "A1:C3").
            table_name: The name of the new table.
        """
        ...

    def agent_summary(self) -> str:
        """
        Return a summary of the workbook's contents, including sheets and tables.

        Returns:
            A string summarizing the workbook's sheets and tables.
        """
        ...
