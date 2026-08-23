from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace

import pytest
from indemnipy_ai.capabilities.excel import (
    DateParsingOptions,
    ExcelCapability,
    ExcelDeps,
    ExcelRuntimeState,
)
from inline_snapshot import snapshot
from pydantic_ai._run_context import RunContext
from pydantic_ai.models.test import TestModel
from pydantic_ai.usage import RunUsage

TEST_DATA_DIR = Path(__file__).parent.parent.parent / "data"
XLSX = TEST_DATA_DIR / "Vantris_Pharmaceuticals_STP_Submission_2026.xlsx"


def _make_ctx(deps):
    return RunContext(deps=deps, model=TestModel(), usage=RunUsage())


@dataclass
class _DepsWithState:
    excel_runtime_state: ExcelRuntimeState


class _DepsWithoutState:
    pass


# ---------------------------------------------------------------------------
# ExcelCapability defaults
# ---------------------------------------------------------------------------


def test_excel_capability_defaults():
    cap = ExcelCapability()
    assert cap.id == snapshot("indemnipy-ai.capabilities.excel")
    assert cap.description == snapshot(
        "Provides tools for reading and analysing spreadsheets."
    )
    assert cap.runtime_state == ExcelRuntimeState()
    assert cap.date_parsing_options == snapshot(
        DateParsingOptions(parse_dates=True, relaxed_about_day=True)
    )


# ---------------------------------------------------------------------------
# get_description
# ---------------------------------------------------------------------------


def test_get_description_returns_description_field():
    assert ExcelCapability(description="custom desc").get_description() == snapshot(
        "custom desc"
    )


def test_get_description_returns_none_when_description_is_none():
    assert ExcelCapability(description=None).get_description() is None


# ---------------------------------------------------------------------------
# get_instructions
# ---------------------------------------------------------------------------


def test_get_instructions_returns_nonempty_string():
    result = ExcelCapability().get_instructions()
    assert isinstance(result, str)
    assert len(result) > 0


def test_get_instructions_is_cached():
    from indemnipy_ai.capabilities.excel._capability import (
        _load_capability_instructions,
    )

    assert _load_capability_instructions() is _load_capability_instructions()


# ---------------------------------------------------------------------------
# get_toolset
# ---------------------------------------------------------------------------


def test_get_toolset_registers_expected_tool_names():
    toolset = ExcelCapability().get_toolset()
    assert set(toolset.tools.keys()) == {
        "load_workbook",
        "list_workbooks",
        "get_workbook_vba",
        "list_tables_and_metadata",
        "preview_table",
        "query_table",
        "list_worksheets",
        "get_range",
        "add_table_from_range",
        "query_store_and_preview",
        "list_derived_tables",
        "preview_derived_table",
    }


def test_get_toolset_preloads_workbooks_listed_in_runtime_state_spreadsheets():
    state = ExcelRuntimeState(spreadsheets=[XLSX])
    cap = ExcelCapability(runtime_state=state)
    cap.get_toolset()
    assert XLSX.name in state.workbooks


# ---------------------------------------------------------------------------
# for_run
# ---------------------------------------------------------------------------


@pytest.mark.anyio
async def test_for_run_reuses_runtime_state_when_deps_implements_excel_deps():
    shared_state = ExcelRuntimeState()
    deps = _DepsWithState(excel_runtime_state=shared_state)
    cap = ExcelCapability()
    result = await cap.for_run(_make_ctx(deps))
    assert result.runtime_state is shared_state


@pytest.mark.anyio
async def test_for_run_creates_fresh_state_when_deps_does_not_implement_excel_deps():
    cap = ExcelCapability()
    original_state = cap.runtime_state
    result = await cap.for_run(_make_ctx(_DepsWithoutState()))
    assert result.runtime_state is not original_state
    assert result.runtime_state == ExcelRuntimeState()


@pytest.mark.anyio
async def test_for_run_preserves_id_description_and_date_options():
    opts = DateParsingOptions(parse_dates=False, relaxed_about_day=False)
    cap = ExcelCapability(
        id="custom-id",
        description="custom-desc",
        date_parsing_options=opts,
    )
    result = await cap.for_run(_make_ctx(_DepsWithoutState()))
    assert result.id == "custom-id"
    assert result.description == "custom-desc"
    assert result.date_parsing_options == opts


# ---------------------------------------------------------------------------
# ExcelDeps protocol
# ---------------------------------------------------------------------------


def test_excel_deps_protocol_satisfied_by_object_with_correct_attribute():
    assert isinstance(
        _DepsWithState(excel_runtime_state=ExcelRuntimeState()), ExcelDeps
    )


def test_excel_deps_protocol_not_satisfied_without_attribute():
    assert not isinstance(_DepsWithoutState(), ExcelDeps)
    assert not isinstance(object(), ExcelDeps)
    assert not isinstance(SimpleNamespace(other_attr=1), ExcelDeps)
