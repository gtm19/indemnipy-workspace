from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, override

from pydantic_ai import RunContext, Tool
from pydantic_ai.capabilities import AbstractCapability
from pydantic_ai.toolsets import FunctionToolset

from indemnipy_ai.capabilities.excel._capability_state import ExcelRuntimeState
from indemnipy_ai.capabilities.excel._models import ExcelWorkbook

from ._toolset import ExcelToolset


@dataclass
class ExcelCapability(AbstractCapability[Any]):
    """
    Capability for working with Excel files.
    """

    id: str | None = "indemnipy-ai.capabilities.excel"

    runtime_state: "ExcelRuntimeState" = field(
        default_factory=lambda: ExcelRuntimeState()
    )

    @override
    async def for_run(self, ctx: RunContext[Any]) -> "ExcelCapability":
        """
        Prepare the capability for a run.

        Args:
            ctx (RunContext[Any]): The run context.
        """
        return ExcelCapability(
            id=self.id,
            description=self.description,
            runtime_state=ExcelRuntimeState(),
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

        return ExcelToolset(
            id=self.id or "indemnipy-ai.capabilities.excel",
            runtime_state=self.runtime_state,
        ).toolset()
