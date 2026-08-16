import logging
from collections.abc import Callable
from dataclasses import dataclass
from pathlib import Path
from typing import Any, NamedTuple, TypeAlias, TypedDict

import duckdb
from pydantic_ai import ModelRetry, Tool
from pydantic_ai.toolsets import FunctionToolset

from indemnipy_ai.capabilities.excel._capability_state import ExcelRuntimeState
from indemnipy_ai.capabilities.excel._functions import DateParsingOptions, _print_df
from indemnipy_ai.capabilities.excel._models import _ExcelWorkbook
from indemnipy_ai.capabilities.excel._workbook_protocol import (
    WorkbookSheet,
    WorkbookTable,
)

# {col_name: dtype_string}
_ColumnSchema: TypeAlias = dict[str, str]
# {table_name: _ColumnSchema}
_TableSchema: TypeAlias = dict[str, _ColumnSchema]
# {sheet_name: _TableSchema}  (only sheets with at least one table)
_SheetSchema: TypeAlias = dict[str, _TableSchema]
# {workbook_name: _SheetSchema}
_WorkbookSchema: TypeAlias = dict[str, _SheetSchema]


class _SheetInfo(TypedDict):
    name: str
    range: str
    freeze_panes: str | None
    min_column: int
    min_row: int
    max_column: int
    max_row: int
    state: str
    tables: list[str]


@dataclass
class _QueryResult:
    table_name: str
    row_count: int
    preview: str


logger = logging.getLogger(__name__)


class _QueryTable(NamedTuple):
    table: str
    """The name of the table."""
    reference_name: str
    """The alias used to reference this table in the SQL query."""
    workbook: str | None = None
    """The workbook containing the table. Omit to reference a derived table."""
    sheet: str | None = None
    """The sheet containing the table. Required when workbook is specified."""


@dataclass
class _ExcelToolset:
    id: str
    runtime_state: "ExcelRuntimeState"
    date_parsing_options: "DateParsingOptions"

    def _get_workbook_table(
        self, workbook_name: str, sheet_name: str, table_name: str
    ) -> WorkbookTable | None:
        workbook = self.runtime_state.workbooks.get(workbook_name)
        if not workbook:
            return None
        sheet = next((s for s in workbook.sheets if s.name == sheet_name), None)
        if not sheet:
            return None
        return next((t for t in sheet.tables if t.name == table_name), None)

    def load_workbook(self, file_path: Path) -> str:
        """
        Load an Excel workbook and add it to the capability's workbooks list.

        Args:
            file_path (Path): The path to the Excel file.

        Returns:
            The key under which the workbook was stored.
        """
        base = file_path.name
        existing = [
            k
            for k in self.runtime_state.workbooks
            if k == base or k.startswith(base + " (")
        ]
        key = f"{base} ({len(existing)})" if existing else base
        workbook = _ExcelWorkbook(
            filepath=file_path, date_parsing_options=self.date_parsing_options
        )
        self.runtime_state.workbooks[key] = workbook
        return key

    def list_workbooks(self) -> str:
        """
        List information about all workbooks.

        Returns:
            A string representation of the workbooks and their sheets / tables
        """
        summaries = [wb.agent_summary() for wb in self.runtime_state.workbooks.values()]
        return "\n\n".join(summaries) if summaries else "No workbooks loaded."

    def list_worksheets(
        self, workbook_name: str | None = None
    ) -> dict[str, list[_SheetInfo]] | None:
        """
        List the worksheets in a specific loaded Excel workbook, or all workbooks if no name is provided.

        Returns sheet metadata and the names of any tables defined in each sheet.
        Use list_tables_and_metadata() to get full column-level detail for those tables.

        Args:
            workbook_name (str, optional): The name of the workbook to list worksheets from.
                If None, lists worksheets from all loaded workbooks.

        Returns:
            A dictionary mapping workbook names to a list of sheet metadata dicts, or None if the
            workbook does not exist. Each dict contains: name, range, freeze_panes, min_column,
            min_row, max_column, max_row, state, and tables (list of table names only).
        """

        def _sheet_to_dict(sheet: WorkbookSheet) -> _SheetInfo:
            return {
                "name": sheet.name,
                "range": sheet.range,
                "freeze_panes": sheet.freeze_panes,
                "min_column": sheet.min_column,
                "min_row": sheet.min_row,
                "max_column": sheet.max_column,
                "max_row": sheet.max_row,
                "state": sheet.state,
                "tables": [t.name for t in sheet.tables],
            }

        if workbook_name:
            workbook = self.runtime_state.workbooks.get(workbook_name)
            if workbook:
                return {workbook_name: [_sheet_to_dict(s) for s in workbook.sheets]}
            return None
        return {
            name: [_sheet_to_dict(s) for s in wb.sheets]
            for name, wb in self.runtime_state.workbooks.items()
        }

    def get_range(
        self, workbook_name: str, sheet_name: str, range_str: str = "A1:J100"
    ) -> list[list[Any]] | None:
        """
        Get the values in a specified range from a given sheet in a loaded Excel workbook.

        Args:
            workbook_name (str): The name of the workbook.
            sheet_name (str): The name of the sheet to retrieve the range from.
            range_str (str): The range string in Excel format (e.g., "A1:C10"). Default is "A1:J100".

        Returns:
            list[list[Any]] | None: A list of lists containing the values in the specified range,
                or None if the workbook or sheet does not exist.
        """
        workbook = self.runtime_state.workbooks.get(workbook_name)
        if not workbook:
            return None

        try:
            return workbook.get_range(sheet_name, range_str)
        except Exception as e:
            raise ModelRetry(
                f"Error retrieving range '{range_str}' from sheet '{sheet_name}' in workbook '{workbook_name}': {e}"
            ) from e

    def list_tables_and_metadata(
        self, workbook_name: str | None = None
    ) -> _WorkbookSchema:
        """
        List the tables in loaded Excel workbooks, organised by workbook and sheet, along with their column names and types.

        Only sheets that contain at least one table are included.

        Args:
            workbook_name (str | None): The name of the workbook to list tables from.
                If None, lists tables from all loaded workbooks.

        Returns:
            A nested dictionary: {workbook_name: {sheet_name: {table_name: {col: dtype}}}}
        """

        def _sheet_tables(sheet: WorkbookSheet) -> _TableSchema:
            return {
                table.name: {
                    col: str(table.dataframe[col].dtype)
                    for col in table.dataframe.columns
                }
                for table in sheet.tables
            }

        if workbook_name:
            workbook = self.runtime_state.workbooks.get(workbook_name)
            if workbook:
                return {
                    workbook_name: {
                        sheet.name: _sheet_tables(sheet)
                        for sheet in workbook.sheets
                        if sheet.tables
                    }
                }
            return {}
        return {
            name: {
                sheet.name: _sheet_tables(sheet) for sheet in wb.sheets if sheet.tables
            }
            for name, wb in self.runtime_state.workbooks.items()
        }

    def preview_table(
        self,
        workbook_name: str,
        sheet_name: str,
        table_name: str,
        n_rows: int = 10,
        offset: int = 0,
    ) -> str | None:
        """
        Preview the first few rows of a specific table in a loaded Excel workbook.

        Args:
            workbook_name (str): The name of the workbook.
            sheet_name (str): The name of the sheet containing the table.
            table_name (str): The name of the table to preview.
            n_rows (int): The number of rows to preview. Default is 10.
            offset (int): The number of rows to skip before starting the preview. Default is 0.

        Returns:
            str | None: The previewed rows as a string representation of a Polars DataFrame, or None if the workbook,
                sheet, or table does not exist.
        """
        table = self._get_workbook_table(workbook_name, sheet_name, table_name)
        if not table:
            return None
        return _print_df(table.dataframe[offset : offset + n_rows], hide_shape=True)

    def add_table_from_range(
        self, workbook_name: str, sheet_name: str, range_str: str, table_name: str
    ) -> _TableSchema | None:
        """
        Add a new table to a workbook sheet from a cell range. The table name must not already
        exist in that sheet — use a unique name. To transform or combine table data, use query()
        instead and store the result as a derived table.

        Args:
            workbook_name: The name of the workbook.
            sheet_name: The name of the sheet containing the range.
            range_str: The Excel-style range string for the new table (e.g., "A1:C3").
            table_name: The name to assign to the new table. Must be unique within the sheet.

        Returns:
            dict | None: The column schema of the added table {table_name: {col: (dtype,)}},
                matching the format of list_tables_and_metadata(), or None if the workbook does not exist.
        """
        workbook = self.runtime_state.workbooks.get(workbook_name)
        if not workbook:
            return None

        if self._get_workbook_table(workbook_name, sheet_name, table_name) is not None:
            raise ModelRetry(
                f"Table '{table_name}' already exists in sheet '{sheet_name}' of workbook '{workbook_name}'. "
                + "Choose a different name, or use query() to transform the data into a derived table."
            )

        try:
            workbook.add_table_from_range(sheet_name, range_str, table_name)
            table = self._get_workbook_table(workbook_name, sheet_name, table_name)
            if table is None:
                return None
            return {
                table_name: {
                    col: str(table.dataframe[col].dtype)
                    for col in table.dataframe.columns
                }
            }
        except Exception as e:
            raise ModelRetry(
                f"Error adding table '{table_name}' from range '{range_str}' in sheet '{sheet_name}' of workbook '{workbook_name}': {e}"
            ) from e

    def query_table(
        self, workbook_name: str, sheet_name: str, table_name: str, query: str
    ) -> str | None:
        """
        Execute a read-only SQL query on a specific workbook table, returning results without modifying the table.

        Args:
            workbook_name (str): The name of the workbook.
            sheet_name (str): The name of the sheet containing the table.
            table_name (str): The name of the table to query.
            query (str): A DuckDB SQL query referencing the table as `df`.

        Returns:
            str | None: The query results as a string (formatted DataFrame), or None if not found.
        """
        table = self._get_workbook_table(workbook_name, sheet_name, table_name)
        if not table:
            return None

        try:
            with duckdb.connect() as con:
                con.register("df", table.dataframe)
                return _print_df(con.sql(query).pl(), hide_shape=True)
        except Exception as e:
            raise ModelRetry(
                f"Error querying table '{table_name}' in sheet '{sheet_name}' of workbook '{workbook_name}': {e}"
            ) from e

    def get_workbook_vba(self, workbook_name: str) -> str | None:
        """
        Get the VBA summary of a loaded Excel workbook by name.

        Args:
            workbook_name (str): The name of the workbook.

        Returns:
            str | None: The VBA summary in Markdown format, or None if not found.
        """
        workbook = self.runtime_state.workbooks.get(workbook_name)
        if workbook and workbook.vba_summary:
            return workbook.vba_summary.to_md()
        return None

    def query_store_and_preview(
        self,
        tables_used: list[_QueryTable],
        query: str,
        table_name: str,
        preview_rows: int = 5,
    ) -> _QueryResult | None:
        """
        Execute a SQL query across multiple tables and store the results as a derived table, and return a preview of the results.

        Tables can come from workbook sheets or from previously created derived tables.
        Results are stored in the session's derived table store and are accessible via
        list_derived_tables() and preview_derived_table().

        Args:
            tables_used (list[_QueryTable]): Tables to include in the query. Each entry specifies
                the table name and SQL alias, and optionally the workbook and sheet. If workbook
                and sheet are omitted the table is looked up from derived tables.
            query (str): DuckDB SQL query referencing each table by its reference_name alias.
            table_name (str): The name to store the results under in the derived table store.

        Returns:
            dict | None: {"table_name": str, "row_count": int, "preview": list[dict]} — the name
                of the stored derived table, its total row count, and a preview of the first 5 rows.
                Returns None on failure.
        """
        dataframes: dict[str, WorkbookTable] = {}

        for qt in tables_used:
            if qt.workbook is not None:
                if qt.sheet is None:
                    raise ModelRetry(
                        f"sheet is required when workbook is specified (table '{qt.table}')."
                    )
                table = self._get_workbook_table(qt.workbook, qt.sheet, qt.table)
                if table is None:
                    raise ModelRetry(
                        f"Table '{qt.table}' not found in sheet '{qt.sheet}' of workbook '{qt.workbook}'."
                    )
            else:
                table = self.runtime_state.derived_tables.get(qt.table)
                if table is None:
                    raise ModelRetry(f"Derived table '{qt.table}' not found.")
            dataframes[qt.reference_name] = table

        try:
            with duckdb.connect() as con:
                for alias, tbl in dataframes.items():
                    con.register(alias, tbl.dataframe)
                results = con.sql(query).pl()
            logger.debug(f"Table:\n\n{results}")
            self.runtime_state.derived_tables[table_name] = WorkbookTable(
                name=table_name,
                sheet_name="",
                range="",
                dataframe=results,
            )
            return _QueryResult(
                table_name=table_name,
                row_count=len(results),
                preview=_print_df(results.head(preview_rows), hide_shape=True),
            )
        except ModelRetry:
            raise
        except Exception as e:
            raise ModelRetry(
                f"Error executing query on tables {[(qt.workbook, qt.sheet, qt.table) for qt in tables_used]}: {e}"
            ) from e

    def list_derived_tables(self) -> _TableSchema:
        """
        List all derived tables (and their column names and types) created during this session.

        Derived tables are the results of query() calls and are not tied to any workbook
        or worksheet. They can be used as inputs to further query() calls or inspected
        with preview_derived_table().

        Returns:
            A dictionary mapping table names to their column names and types.
        """
        return {
            name: {
                col: str(table.dataframe[col].dtype) for col in table.dataframe.columns
            }
            for name, table in self.runtime_state.derived_tables.items()
        }

    def preview_derived_table(
        self, table_name: str, n_rows: int = 10, offset: int = 0
    ) -> str | None:
        """
        Preview the first few rows of a derived table.

        Args:
            table_name (str): The name of the derived table.
            n_rows (int): The number of rows to preview. Default is 10.

        Returns:
            str | None: The previewed rows as a string (formatted DataFrame),
                or None if the table does not exist.
        """
        table = self.runtime_state.derived_tables.get(table_name)
        if not table:
            return None
        return _print_df(table.dataframe[offset : offset + n_rows], hide_shape=True)

    def toolset(self) -> FunctionToolset:
        functions: set[Callable[..., Any]] = {
            self.load_workbook,
            self.list_workbooks,
            self.get_workbook_vba,
            self.list_tables_and_metadata,
            self.preview_table,
            self.query_table,
            self.list_worksheets,
            self.get_range,
            self.add_table_from_range,
            self.query_store_and_preview,
            self.list_derived_tables,
            self.preview_derived_table,
        }

        excel_toolset = FunctionToolset(
            tools=[
                Tool(
                    fn,
                    name=fn.__name__,
                    description=fn.__doc__ or "",
                )
                for fn in functions
            ],
            id=self.id,
        )

        return excel_toolset
