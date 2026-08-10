from dataclasses import dataclass
from pathlib import Path
from typing import Protocol


@dataclass(frozen=True)
class ParsedVbaAnalysisResult:
    kw_type: str
    keyword: str
    description: str


@dataclass(frozen=True)
class ParsedVbaMacro:
    filename: str
    stream_path: str
    vba_filename: str
    vba_code: str


@dataclass(frozen=True)
class VbaParseResult:
    analysis_results: tuple[ParsedVbaAnalysisResult, ...]
    macros: tuple[ParsedVbaMacro, ...]


class VbaParser(Protocol):
    def __call__(self, filepath: Path) -> VbaParseResult: ...
