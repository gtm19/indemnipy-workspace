from dataclasses import dataclass, field

from indemnipy_ai.capabilities.excel import (
    ExcelCapability,
    ExcelRuntimeState,
)
from pydantic_ai import Agent


@dataclass
class Deps:
    excel_runtime_state: ExcelRuntimeState = field(default_factory=ExcelRuntimeState)


agent = Agent(
    "test",
    capabilities=[ExcelCapability()],
    output_type=str,
    deps_type=Deps,
)
deps = Deps(excel_runtime_state=ExcelRuntimeState(excel_files=[]))
print("OK")
