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

    spreadsheets: list[Path] = field(default_factory=list)
    """List of spreadsheets that have been loaded during the agent's execution."""
    workbooks: dict[str, "WorkbookProtocol"] = field(default_factory=dict)
    """Mapping of file paths to loaded workbook instances."""
    derived_tables: dict[str, "WorkbookTable"] = field(default_factory=dict)
    """Mapping of derived table names to their corresponding WorkbookTable instances."""
