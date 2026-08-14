from dataclasses import dataclass
from pathlib import Path
from typing import Protocol

import polars as pl


# VBA Parser
@dataclass(frozen=True)
class _VbaAnalysisResult:
    kw_type: str
    keyword: str
    description: str


@dataclass(frozen=True)
class _VbaMacro:
    filename: str
    stream_path: str
    vba_filename: str
    vba_code: str


@dataclass(frozen=True)
class _VbaParseResult:
    analysis_results: tuple[_VbaAnalysisResult, ...]
    macros: tuple[_VbaMacro, ...]


class _VbaParser(Protocol):
    def __call__(self, filepath: Path) -> _VbaParseResult: ...


# Table Parser
@dataclass(frozen=True)
class _ParsedWorkbookTable:
    name: str
    sheet_name: str
    range: str
    dataframe: "pl.DataFrame"
