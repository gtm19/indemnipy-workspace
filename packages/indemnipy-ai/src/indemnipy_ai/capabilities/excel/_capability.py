import functools
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Protocol, override, runtime_checkable

from pydantic_ai import RunContext
from pydantic_ai.capabilities import AbstractCapability
from pydantic_ai.toolsets import FunctionToolset

from indemnipy_ai.capabilities.excel._capability_state import ExcelRuntimeState
from indemnipy_ai.capabilities.excel._functions import DateParsingOptions

from ._toolset import _ExcelToolset

CURRENT_DIR = Path(__file__).parent


@functools.cache
def _load_capability_instructions() -> str:
    return (CURRENT_DIR / "_capability_instructions.md").read_text()


@runtime_checkable
class ExcelDeps(Protocol):
    """
    Dependency protocol for Excel capabilities.

    Pass the same instance across multiple agent runs to preserve loaded workbooks
    and derived tables between turns in a multi-turn conversation.

    Attributes:
        excel_runtime_state: Mutable state holding loaded workbooks and derived tables.
    """

    excel_runtime_state: ExcelRuntimeState


@dataclass
class ExcelCapability(AbstractCapability[ExcelDeps | Any]):
    """
    Capability for working with Excel files.
    """

    id: str | None = "indemnipy-ai.capabilities.excel"

    runtime_state: "ExcelRuntimeState" = field(default_factory=ExcelRuntimeState)
    date_parsing_options: DateParsingOptions = field(
        default_factory=lambda: DateParsingOptions(
            parse_dates=True, relaxed_about_day=True
        )
    )

    @override
    async def for_run(self, ctx: RunContext[ExcelDeps | Any]) -> "ExcelCapability":
        """
        Prepare the capability for a run.

        If deps implements ExcelDeps, the existing runtime state is reused so that
        workbooks and derived tables loaded in a previous turn are still available.
        Otherwise a fresh ExcelRuntimeState is created.

        Args:
            ctx (RunContext[ExcelDeps | Any]): The run context.
        """
        runtime_state = (
            ctx.deps.excel_runtime_state
            if isinstance(ctx.deps, ExcelDeps)
            else ExcelRuntimeState()
        )
        return ExcelCapability(
            id=self.id,
            description=self.description,
            runtime_state=runtime_state,
            date_parsing_options=self.date_parsing_options,
        )

    @override
    def get_description(self) -> str:
        return "Capability for working with Excel files." + (self.description or "")

    @override
    def get_toolset(self) -> FunctionToolset:
        """
        Get the toolset for Excel capabilities.

        Returns:
            FunctionToolset: The toolset containing Excel-related tools.
        """
        excel_toolset_builder = _ExcelToolset(
            id=self.id or "indemnipy-ai.capabilities.excel",
            runtime_state=self.runtime_state,
            date_parsing_options=self.date_parsing_options,
        )

        if self.runtime_state.excel_files:
            for file_path in self.runtime_state.excel_files:
                _ = excel_toolset_builder.load_workbook(file_path)
        return excel_toolset_builder.toolset()

    @override
    def get_instructions(self) -> str:
        return _load_capability_instructions()
