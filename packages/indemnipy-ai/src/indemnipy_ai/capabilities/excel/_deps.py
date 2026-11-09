from typing import Protocol, runtime_checkable

from indemnipy_ai.capabilities.excel._capability_state import ExcelRuntimeState


@runtime_checkable
class ExcelDeps(Protocol):
    """Dependency protocol for spreadsheet capabilities.

    Any object that has an ``excel_runtime_state`` attribute of type
    [ExcelRuntimeState][indemnipy_ai.capabilities.excel.ExcelRuntimeState]
    satisfies this protocol — no explicit inheritance is required.

    Pass the same instance across multiple agent runs to preserve loaded
    workbooks and derived tables between turns in a multi-turn conversation.

    Attributes:
        excel_runtime_state: Mutable state holding loaded workbooks and derived
            tables.
    """

    excel_runtime_state: ExcelRuntimeState
