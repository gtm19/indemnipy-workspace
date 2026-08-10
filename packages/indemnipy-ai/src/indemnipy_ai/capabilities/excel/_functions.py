from pathlib import Path
from typing import cast

from oletools.olevba import VBA_Parser  # pyright: ignore[reportMissingTypeStubs]

from ._parser_contract import (
    ParsedVbaAnalysisResult,
    ParsedVbaMacro,
    VbaParseResult,
)


def oletools_vba_parser(filepath: Path) -> VbaParseResult:
    parser = VBA_Parser(str(filepath))

    analysis_results: list[ParsedVbaAnalysisResult] = []
    if parser.detect_vba_macros():
        analysis_results_raw = parser.analyze_macros()
        if analysis_results_raw is not None:
            for kw_type, keyword, description in analysis_results_raw:
                analysis_results.append(
                    ParsedVbaAnalysisResult(
                        kw_type=kw_type,
                        keyword=keyword,
                        description=description,
                    )
                )

    macros: list[ParsedVbaMacro] = []
    for filename, stream_path, vba_filename, vba_code in parser.extract_macros():
        macros.append(
            ParsedVbaMacro(
                filename=cast(str, filename),
                stream_path=cast(str, stream_path),
                vba_filename=cast(str, vba_filename),
                vba_code=cast(str, vba_code),
            )
        )

    return VbaParseResult(
        analysis_results=tuple(analysis_results),
        macros=tuple(macros),
    )
