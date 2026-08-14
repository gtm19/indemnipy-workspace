from dataclasses import dataclass, field
from pathlib import Path

from indemnipy_ai.capabilities.excel._workbook_protocol import (
    WorkbookProtocol,
    WorkbookTable,
)


@dataclass
class ExcelRuntimeState:
    """
    Runtime state for ExcelCapability. Holds loaded workbooks and derived tables.

    Pass the same instance across multiple agent runs to preserve loaded workbooks
    and derived tables between turns in a multi-turn conversation.
    """

    excel_files: list[Path] = field(default_factory=list)
    workbooks: dict[str, "WorkbookProtocol"] = field(default_factory=dict)
    derived_tables: dict[str, "WorkbookTable"] = field(default_factory=dict)
