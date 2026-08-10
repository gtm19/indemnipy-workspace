from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from pydantic_ai import Tool
from pydantic_ai.toolsets import FunctionToolset

from indemnipy_ai.capabilities.excel._capability_state import ExcelRuntimeState
from indemnipy_ai.capabilities.excel._models import ExcelWorkbook


@dataclass
class ExcelToolset:
    id: str
    runtime_state: "ExcelRuntimeState"

    def load_workbook(self, file_path: Path) -> None:
        """
        Load an Excel workbook and add it to the capability's workbooks list.

        Args:
            file_path (Path): The path to the Excel file.
        """
        if file_path.name in self.runtime_state.workbooks:
            n_occ = sum(
                1 for name in self.runtime_state.workbooks if name == file_path.name
            )
            key = file_path.name + f" ({n_occ})"
        else:
            key = file_path.name
        workbook = ExcelWorkbook.from_file(file_path)
        self.runtime_state.workbooks[key] = workbook

    def list_workbooks(self) -> list[str]:
        """
        List the names of all loaded Excel workbooks.

        Returns:
            list[str]: A list of workbook names.
        """
        return list(self.runtime_state.workbooks.keys())

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

    def toolset(self) -> FunctionToolset:
        functions: set[Callable[..., Any]] = {
            self.load_workbook,
            self.list_workbooks,
            self.get_workbook_vba,
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
