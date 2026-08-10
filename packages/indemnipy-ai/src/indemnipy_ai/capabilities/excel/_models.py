from dataclasses import dataclass
from pathlib import Path

from indemnipy_ai.capabilities.excel._functions import (
    oletools_vba_parser,
)
from indemnipy_ai.capabilities.excel._parser_contract import VbaParser


@dataclass
class VbaMacro:
    """Represents a single extracted VBA macro stream.

    Attributes:
        filename: Container file that holds the VBA project data.
        stream_path: Path of the stream within the OLE or workbook structure.
        vba_filename: VBA module or class filename.
        vba_code: Raw VBA source code extracted from the stream.
    """

    filename: str
    stream_path: str
    vba_filename: str
    vba_code: str


@dataclass
class VbaAnalysisResult:
    """Represents one heuristic analysis result.

    Attributes:
        kw_type: Category of the finding, such as `Suspicious`.
        keyword: Keyword or pattern that triggered the finding.
        description: Human-readable explanation of the finding.
    """

    kw_type: str
    keyword: str
    description: str


@dataclass
class VbaSummary:
    """Aggregates extracted VBA macros and parser analysis for one file.

    Attributes:
        filepath: Workbook path that was inspected.
        analysis_results: Heuristic findings returned by macro analysis.
        macros: Extracted VBA macro streams.
    """

    filepath: Path
    analysis_results: list[VbaAnalysisResult]
    macros: list[VbaMacro]

    @classmethod
    def from_file(
        cls, filepath: Path, parser: VbaParser = oletools_vba_parser
    ) -> "VbaSummary":
        """Build a summary from an Excel workbook or OLE document.

        Args:
            filepath: Path to the workbook to inspect.

        Returns:
            A summary containing any detected analysis findings and extracted
            VBA macro streams.
        """

        parse_result = parser(filepath)
        return cls(
            filepath=filepath,
            analysis_results=[
                VbaAnalysisResult(
                    kw_type=r.kw_type, keyword=r.keyword, description=r.description
                )
                for r in parse_result.analysis_results
            ],
            macros=[
                VbaMacro(
                    filename=m.filename,
                    stream_path=m.stream_path,
                    vba_filename=m.vba_filename,
                    vba_code=m.vba_code,
                )
                for m in parse_result.macros
            ],
        )

    def to_md(self) -> str:
        """Render the summary as Markdown.

        Returns:
            A Markdown document containing analysis results followed by each
            extracted macro and its VBA source code.
        """

        content: str = f"# VBA Macros in {self.filepath.name}\n"

        if self.analysis_results:
            content += f"\n## Analysis Results\nThere are {len(self.analysis_results)} observations.\n"
            for i, result in enumerate(self.analysis_results):
                content += f"{i + 1}. Type: {result.kw_type}, Keyword: {result.keyword}, Description: {result.description}\n"

        for i, macro in enumerate(self.macros):
            content += f"\n## Macro {i + 1}\n"
            content += f"Filename: {macro.filename}\n"
            content += f"Stream Path: {macro.stream_path}\n"
            content += f"VBA Filename: {macro.vba_filename}\n\n"
            content += f"""### VBA Code:
```vba
{macro.vba_code}
```
"""
        return content


@dataclass
class ExcelWorkbook:
    """Represents an Excel workbook file.

    Attributes:
        filepath: Path to the workbook file.
        vba_summary: Optional summary of any detected VBA macros and analysis results.
    """

    filepath: Path
    vba_summary: VbaSummary | None = None

    @classmethod
    def from_file(cls, filepath: Path) -> "ExcelWorkbook":
        """Build an ExcelWorkbook instance from a file.

        Args:
            filepath: Path to the workbook to inspect.

        Returns:
            An ExcelWorkbook instance with an optional VBA summary if macros are present.
        """
        vba_summary = VbaSummary.from_file(filepath)
        if not vba_summary.macros and not vba_summary.analysis_results:
            vba_summary = None
        return cls(filepath=filepath, vba_summary=vba_summary)
