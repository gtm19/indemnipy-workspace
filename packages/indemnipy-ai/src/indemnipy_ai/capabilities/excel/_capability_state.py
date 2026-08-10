from dataclasses import dataclass, field

from indemnipy_ai.capabilities.excel._models import ExcelWorkbook


@dataclass
class ExcelRuntimeState:
    """
    Represents the runtime state of the Excel capability.
    """

    workbooks: dict[str, "ExcelWorkbook"] = field(default_factory=dict)
