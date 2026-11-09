# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.1.4] - 2026-11-08

### Added

- `ExcelCapability` now integrates with Pydantic AI's workspace system. Set `base_dir` to the directory containing your spreadsheets; the capability exposes a read-only workspace rooted there so the agent can browse for files.
- New `list_available_workbooks` tool — lets the agent discover `.xlsx` and `.xlsm` files within the workspace without requiring the user to supply explicit paths.

### Changed

- `load_workbook` is now async and workspace-aware. It validates that the requested file is within the configured workspace root and rejects paths that escape it.
- `list_workbooks` renamed to `list_loaded_workbooks` for clarity.
- `list_worksheets` renamed to `list_worksheets_from_loaded_workbook` for clarity.

## [0.1.3] - 2026-10-28

### Added

- Comprehensive test suite for the Excel toolset
- New use case added to documentation

### Changed

- Documentation fixes and improvements
- Re-added GitHub release workflow to CI (in draft)
- Updated pre-commit hooks to latest versions

### Security

- Added DuckDB configuration to disable external file access in SQL query tools

## [0.1.2] - 2026-08-16

### Fixed

- Broken link in `pyproject.toml`

## [0.1.1] - 2026-08-16

### Added

- CODEOWNERS file
- This changelog

### Changed

- Updated license
- Moved GitHub release to manual

## [0.1.0] - 2026-08-14

### Added

- Initial release
