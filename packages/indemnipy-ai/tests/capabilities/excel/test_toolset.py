from pathlib import Path

import pytest
from indemnipy_ai.capabilities.excel._capability_state import ExcelRuntimeState
from indemnipy_ai.capabilities.excel._functions import DateParsingOptions
from indemnipy_ai.capabilities.excel._toolset import _ExcelToolset, _QueryTable
from inline_snapshot import snapshot
from pydantic_ai import RunContext, RunUsage, ToolFailed
from pydantic_ai.exceptions import ModelRetry
from pydantic_ai.models.test import TestModel
from pydantic_ai.workspaces import (
    LocalWorkspaceBackend,
    ReadOnlyWorkspace,
    Workspace,
)

pytestmark = pytest.mark.anyio

TEST_DATA_DIR = Path(__file__).parent.parent.parent / "data"
XLSX_KEY = "Vantris_Pharmaceuticals_STP_Submission_2026.xlsx"
XLSX = TEST_DATA_DIR / XLSX_KEY
XLSM_KEY = "Vantris_Pharmaceuticals_STP_Submission_2026.xlsm"
XLSM = TEST_DATA_DIR / XLSM_KEY

DEPS = ExcelRuntimeState()
CTX = RunContext(
    deps=DEPS,
    model=TestModel(),
    workspace=ReadOnlyWorkspace(Workspace(LocalWorkspaceBackend(TEST_DATA_DIR))),
    usage=RunUsage(),
)

LOC_TABLE = _QueryTable(
    table="tblLocations", reference_name="loc", workbook=XLSX_KEY, sheet="Locations"
)


async def _make_toolset(path: Path | None = None) -> _ExcelToolset:
    state = ExcelRuntimeState()
    ts = _ExcelToolset(
        id="test",
        runtime_state=state,
        date_parsing_options=DateParsingOptions(
            parse_dates=True, relaxed_about_day=True
        ),
    )
    if path is not None:
        await ts.load_workbook(CTX, path)
    return ts


@pytest.fixture
async def ts_xlsx():
    return await _make_toolset(XLSX)


@pytest.fixture
async def ts_xlsm():
    return await _make_toolset(XLSM)


@pytest.fixture
async def ts_with_uk_locs(ts_xlsx):
    ts_xlsx.query_store_and_preview(
        tables_used=[LOC_TABLE],
        query='SELECT "Location ID", "Country" FROM loc WHERE "Country" = \'United Kingdom\'',
        table_name="uk_locs",
    )
    return ts_xlsx


# ---------------------------------------------------------------------------
# load_workbook
# ---------------------------------------------------------------------------


async def test_load_workbook_returns_filename_as_key():
    ts = await _make_toolset()
    key = await ts.load_workbook(CTX, XLSX)
    assert key == snapshot("Vantris_Pharmaceuticals_STP_Submission_2026.xlsx")
    assert XLSX_KEY in ts.runtime_state.workbooks


async def test_load_workbook_deduplicates_keys_with_numeric_suffix():
    ts = await _make_toolset()
    k1 = await ts.load_workbook(CTX, XLSX)
    k2 = await ts.load_workbook(CTX, XLSX)
    k3 = await ts.load_workbook(CTX, XLSX)
    assert k1 == snapshot("Vantris_Pharmaceuticals_STP_Submission_2026.xlsx")
    assert k2 == snapshot("Vantris_Pharmaceuticals_STP_Submission_2026.xlsx (1)")
    assert k3 == snapshot("Vantris_Pharmaceuticals_STP_Submission_2026.xlsx (2)")


@pytest.mark.parametrize(
    "path",
    [
        # workspace escape via path traversal
        Path("../../evil.xlsx"),
        Path("../../../evil.xlsx"),
        Path("subdir/../../../evil.xlsx"),
        # absolute path outside workspace
        Path("/etc/passwd.xlsx"),
        Path("/tmp/evil.xlsx"),
    ],
    ids=[
        "traversal_two_up",
        "traversal_three_up",
        "traversal_via_subdir",
        "absolute_etc",
        "absolute_tmp",
    ],
)
async def test_load_workbook_rejects_paths_outside_workspace(path):
    ts = await _make_toolset()
    with pytest.raises(ToolFailed):
        await ts.load_workbook(CTX, path)


@pytest.mark.parametrize(
    "path",
    [
        Path("file.csv"),
        Path("file.xls"),  # legacy Excel format
        Path("file.txt"),
        Path("file.zip"),  # xlsx is a zip internally but we don't accept it raw
        Path("file"),  # no extension
    ],
    ids=["csv", "xls", "txt", "zip", "no_extension"],
)
async def test_load_workbook_rejects_unsupported_extensions(path):
    ts = await _make_toolset()
    with pytest.raises(ToolFailed):
        await ts.load_workbook(CTX, path)


# ---------------------------------------------------------------------------
# list_workbooks
# ---------------------------------------------------------------------------


async def test_list_workbooks_empty():
    ts = await _make_toolset()
    assert ts.list_loaded_workbooks() == snapshot("No workbooks loaded.")


async def test_list_workbooks_after_load_contains_filename(ts_xlsx):
    result = ts_xlsx.list_loaded_workbooks()
    assert XLSX_KEY in result
    assert "Workbook:" in result


# ---------------------------------------------------------------------------
# list_worksheets
# ---------------------------------------------------------------------------


async def test_list_worksheets_with_workbook_name(ts_xlsx):
    result = ts_xlsx.list_worksheets_from_loaded_workbook(XLSX_KEY)
    assert result is not None
    assert {k: [s.model_dump() for s in v] for k, v in result.items()} == snapshot(
        {
            "Vantris_Pharmaceuticals_STP_Submission_2026.xlsx": [
                {
                    "name": "Information",
                    "range": "B1:C69",
                    "freeze_panes": None,
                    "min_column": 2,
                    "min_row": 1,
                    "max_column": 3,
                    "max_row": 69,
                    "state": "visible",
                    "tables": [],
                },
                {
                    "name": "Claims",
                    "range": "A1:L39",
                    "freeze_panes": "A11",
                    "min_column": 1,
                    "min_row": 1,
                    "max_column": 12,
                    "max_row": 39,
                    "state": "visible",
                    "tables": [],
                },
                {
                    "name": "Locations",
                    "range": "A1:O33",
                    "freeze_panes": "A2",
                    "min_column": 1,
                    "min_row": 1,
                    "max_column": 15,
                    "max_row": 33,
                    "state": "visible",
                    "tables": ["tblLocations"],
                },
                {
                    "name": "Sendings",
                    "range": "A1:O121",
                    "freeze_panes": "A2",
                    "min_column": 1,
                    "min_row": 1,
                    "max_column": 15,
                    "max_row": 121,
                    "state": "visible",
                    "tables": ["tblSendings"],
                },
            ]
        }
    )


async def test_list_worksheets_without_name_returns_all_workbooks(ts_xlsx):
    result = ts_xlsx.list_worksheets_from_loaded_workbook()
    assert result is not None
    assert XLSX_KEY in result
    assert len(result[XLSX_KEY]) == 4  # Information, Claims, Locations, Sendings


async def test_list_worksheets_returns_none_for_unknown_workbook(ts_xlsx):
    assert ts_xlsx.list_worksheets_from_loaded_workbook("does-not-exist.xlsx") is None


# ---------------------------------------------------------------------------
# list_tables_and_metadata
# ---------------------------------------------------------------------------


async def test_list_tables_and_metadata(ts_xlsx):
    result = ts_xlsx.list_tables_and_metadata()
    assert result == snapshot(
        {
            "Vantris_Pharmaceuticals_STP_Submission_2026.xlsx": {
                "Locations": {
                    "tblLocations": {
                        "Location ID": "String",
                        "Site Name": "String",
                        "Address": "String",
                        "City": "String",
                        "Country": "String",
                        "Region": "String",
                        "Occupancy": "String",
                        "Ownership": "String",
                        "Temperature Regime": "String",
                        "Sprinklered": "String",
                        "Maximum Stock Value (USD)": "Int64",
                        "Average Stock Value (USD)": "Int64",
                        "Peak Month": "String",
                        "Flood Zone": "String",
                        "Notes": "String",
                    }
                },
                "Sendings": {
                    "tblSendings": {
                        "Sending Reference": "String",
                        "Despatch Date": "Datetime(time_unit='us', time_zone=None)",
                        "Description of Goods": "String",
                        "Product Division": "String",
                        "Packing": "String",
                        "Number of Units": "Int64",
                        "Conveyance": "String",
                        "Carrier / Forwarder": "String",
                        "Origin": "String",
                        "Origin Country": "String",
                        "Destination": "String",
                        "Destination Country": "String",
                        "Incoterms": "String",
                        "Temperature Regime": "String",
                        "Insured Value (USD)": "Int64",
                    }
                },
            }
        }
    )


async def test_list_tables_and_metadata_with_workbook_name_scopes_to_that_workbook(
    ts_xlsx,
):
    result = ts_xlsx.list_tables_and_metadata(XLSX_KEY)
    assert set(result.keys()) == {XLSX_KEY}
    assert "Locations" in result[XLSX_KEY]
    assert "Sendings" in result[XLSX_KEY]


async def test_list_tables_and_metadata_excludes_sheets_with_no_tables(ts_xlsx):
    sheets = ts_xlsx.list_tables_and_metadata()[XLSX_KEY]
    assert "Information" not in sheets
    assert "Claims" not in sheets


async def test_list_tables_and_metadata_returns_empty_dict_for_unknown_workbook(
    ts_xlsx,
):
    assert ts_xlsx.list_tables_and_metadata("no-such.xlsx") == {}


# ---------------------------------------------------------------------------
# preview_table
# ---------------------------------------------------------------------------


async def test_preview_table_returns_string_containing_data(ts_xlsx):
    result = ts_xlsx.preview_table(XLSX_KEY, "Locations", "tblLocations")
    assert isinstance(result, str)
    assert "LOC-001" in result


@pytest.mark.parametrize(
    "workbook,sheet,table",
    [
        ("no.xlsx", "Locations", "tblLocations"),
        (XLSX_KEY, "NoSheet", "tblLocations"),
        (XLSX_KEY, "Locations", "noTable"),
    ],
    ids=["missing_workbook", "missing_sheet", "missing_table"],
)
async def test_preview_table_returns_none_for_bad_args(ts_xlsx, workbook, sheet, table):
    assert ts_xlsx.preview_table(workbook, sheet, table) is None


async def test_preview_table_respects_n_rows(ts_xlsx):
    # tblLocations has 32 data rows; request only 2
    result = ts_xlsx.preview_table(XLSX_KEY, "Locations", "tblLocations", n_rows=2)
    assert result is not None
    assert result.count("LOC-") == 2


async def test_preview_table_respects_offset(ts_xlsx):
    result_from_start = ts_xlsx.preview_table(
        XLSX_KEY, "Locations", "tblLocations", n_rows=1, offset=0
    )
    result_offset = ts_xlsx.preview_table(
        XLSX_KEY, "Locations", "tblLocations", n_rows=1, offset=1
    )
    assert result_from_start is not None
    assert result_offset is not None
    assert result_from_start != result_offset


# ---------------------------------------------------------------------------
# get_range
# ---------------------------------------------------------------------------


async def test_get_range_returns_list_of_lists(ts_xlsx):
    result = ts_xlsx.get_range(XLSX_KEY, "Information", "B1:C3")
    assert result == snapshot(
        [
            ["VANTRIS PHARMACEUTICALS LTD", None],
            [
                "Marine Cargo & Stock Throughput – Renewal Submission for 1 October 2026",
                None,
            ],
            [
                "Prepared by Latterworth Marsden Ltd, 12 Fenchurch Avenue, London EC3M 5BY",
                None,
            ],
        ]
    )


async def test_get_range_returns_none_for_missing_workbook(ts_xlsx):
    assert ts_xlsx.get_range("no.xlsx", "Information", "B1:C3") is None


async def test_get_range_raises_model_retry_for_invalid_range(ts_xlsx):
    with pytest.raises(ModelRetry):
        ts_xlsx.get_range(XLSX_KEY, "Information", "NOT_A_RANGE")


# ---------------------------------------------------------------------------
# add_table_from_range
# ---------------------------------------------------------------------------


async def test_add_table_from_range_success(ts_xlsx):
    result = ts_xlsx.add_table_from_range(XLSX_KEY, "Locations", "A1:C3", "myTable")
    assert result == snapshot(
        {
            "myTable": {
                "Location ID": "String",
                "Site Name": "String",
                "Address": "String",
            }
        }
    )
    assert ts_xlsx.preview_table(XLSX_KEY, "Locations", "myTable") is not None


async def test_add_table_from_range_returns_none_for_missing_workbook(ts_xlsx):
    assert ts_xlsx.add_table_from_range("no.xlsx", "Locations", "A1:C3", "t") is None


async def test_add_table_from_range_raises_model_retry_for_duplicate_table_name(
    ts_xlsx,
):
    with pytest.raises(ModelRetry, match="already exists"):
        ts_xlsx.add_table_from_range(XLSX_KEY, "Locations", "A1:C3", "tblLocations")


# ---------------------------------------------------------------------------
# query_table
# ---------------------------------------------------------------------------


async def test_query_table_returns_formatted_string(ts_xlsx):
    result = ts_xlsx.query_table(
        XLSX_KEY,
        "Locations",
        "tblLocations",
        'SELECT "Location ID", "Site Name", "Country" FROM df LIMIT 3',
    )
    assert result == snapshot(
        " Location ID  Site Name                Country        \n"
        " ---          ---                      ---            \n"
        " str          str                      str            \n"
        " LOC-001      Vantris Basingstoke      United Kingdom \n"
        " LOC-002      Vantris Wrexham DC       United Kingdom \n"
        " LOC-003      Vantris Sterile Ireland  Ireland        "
    )


@pytest.mark.parametrize(
    "workbook,sheet,table",
    [
        ("no.xlsx", "Locations", "tblLocations"),
        (XLSX_KEY, "NoSheet", "tblLocations"),
        (XLSX_KEY, "Locations", "noTable"),
    ],
    ids=["missing_workbook", "missing_sheet", "missing_table"],
)
async def test_query_table_returns_none_for_bad_args(ts_xlsx, workbook, sheet, table):
    assert ts_xlsx.query_table(workbook, sheet, table, "SELECT 1") is None


async def test_query_table_raises_model_retry_for_bad_sql(ts_xlsx):
    with pytest.raises(ModelRetry):
        ts_xlsx.query_table(
            XLSX_KEY, "Locations", "tblLocations", "SELECT FROM WHERE BROKEN"
        )


# ---------------------------------------------------------------------------
# get_workbook_vba
# ---------------------------------------------------------------------------


async def test_get_workbook_vba_returns_markdown_for_xlsm(ts_xlsm):
    result = ts_xlsm.get_workbook_vba(XLSM_KEY)
    assert isinstance(result, str)
    assert result.startswith(
        "# VBA Macros in Vantris_Pharmaceuticals_STP_Submission_2026.xlsm"
    )
    assert "## Analysis Results" in result
    assert "```vba" in result


@pytest.mark.parametrize(
    "key", [XLSX_KEY, "no.xlsm"], ids=["xlsx_no_macros", "unknown_key"]
)
async def test_get_workbook_vba_returns_none(ts_xlsx, key):
    assert ts_xlsx.get_workbook_vba(key) is None


# ---------------------------------------------------------------------------
# query_store_and_preview
# ---------------------------------------------------------------------------


async def test_query_store_and_preview_stores_result_and_returns_query_result(ts_xlsx):
    result = ts_xlsx.query_store_and_preview(
        tables_used=[LOC_TABLE],
        query='SELECT "Location ID", "Country" FROM loc WHERE "Country" = \'United Kingdom\'',
        table_name="uk_locs",
        preview_rows=3,
    )
    assert result is not None
    assert result.table_name == snapshot("uk_locs")
    assert result.row_count == snapshot(3)
    assert result.preview == snapshot(
        " Location ID  Country        \n"
        " ---          ---            \n"
        " str          str            \n"
        " LOC-001      United Kingdom \n"
        " LOC-002      United Kingdom \n"
        " LOC-008      United Kingdom "
    )
    assert "uk_locs" in ts_xlsx.runtime_state.derived_tables


async def test_query_store_and_preview_raises_model_retry_when_sheet_missing_for_workbook_table(
    ts_xlsx,
):
    with pytest.raises(ModelRetry, match="sheet is required"):
        ts_xlsx.query_store_and_preview(
            tables_used=[
                _QueryTable(
                    table="tblLocations",
                    reference_name="t",
                    workbook=XLSX_KEY,
                    sheet=None,
                )
            ],
            query="SELECT * FROM t",
            table_name="out",
        )


async def test_query_store_and_preview_raises_model_retry_for_missing_workbook_table(
    ts_xlsx,
):
    with pytest.raises(ModelRetry, match="not found"):
        ts_xlsx.query_store_and_preview(
            tables_used=[
                _QueryTable(
                    table="noTable",
                    reference_name="t",
                    workbook=XLSX_KEY,
                    sheet="Locations",
                )
            ],
            query="SELECT * FROM t",
            table_name="out",
        )


async def test_query_store_and_preview_raises_model_retry_for_missing_derived_table(
    ts_xlsx,
):
    with pytest.raises(ModelRetry, match="not found"):
        ts_xlsx.query_store_and_preview(
            tables_used=[_QueryTable(table="ghost", reference_name="g")],
            query="SELECT * FROM g",
            table_name="out",
        )


async def test_query_store_and_preview_raises_model_retry_for_bad_sql(ts_xlsx):
    with pytest.raises(ModelRetry):
        ts_xlsx.query_store_and_preview(
            tables_used=[LOC_TABLE],
            query="SELECT FROM WHERE BROKEN",
            table_name="out",
        )


# ---------------------------------------------------------------------------
# list_derived_tables
# ---------------------------------------------------------------------------


async def test_list_derived_tables_is_empty_before_any_queries(ts_xlsx):
    assert ts_xlsx.list_derived_tables() == {}


async def test_list_derived_tables_populated_after_query_store_and_preview(ts_xlsx):
    ts_xlsx.query_store_and_preview(
        tables_used=[LOC_TABLE],
        query='SELECT "Location ID", "Country" FROM loc LIMIT 5',
        table_name="sample",
    )
    assert ts_xlsx.list_derived_tables() == snapshot(
        {"sample": {"Location ID": "String", "Country": "String"}}
    )


# ---------------------------------------------------------------------------
# preview_derived_table
# ---------------------------------------------------------------------------


async def test_preview_derived_table_returns_none_for_unknown_table(ts_xlsx):
    assert ts_xlsx.preview_derived_table("ghost") is None


async def test_preview_derived_table_returns_string_for_known_table(ts_with_uk_locs):
    result = ts_with_uk_locs.preview_derived_table("uk_locs", n_rows=2)
    assert result == snapshot(
        " Location ID  Country        \n"
        " ---          ---            \n"
        " str          str            \n"
        " LOC-001      United Kingdom \n"
        " LOC-002      United Kingdom "
    )


async def test_preview_derived_table_respects_n_rows_and_offset(ts_with_uk_locs):
    r1 = ts_with_uk_locs.preview_derived_table("uk_locs", n_rows=1, offset=0)
    r2 = ts_with_uk_locs.preview_derived_table("uk_locs", n_rows=1, offset=1)
    assert r1 is not None and r2 is not None
    assert r1 != r2


# ---------------------------------------------------------------------------
# DuckDB security: no external file read or write
# ---------------------------------------------------------------------------


async def test_query_table_cannot_read_arbitrary_csv(ts_xlsx, tmp_path):
    csv_file = tmp_path / "secret.csv"
    csv_file.write_text("a,b\n1,2\n3,4\n")
    with pytest.raises(ModelRetry):
        ts_xlsx.query_table(
            XLSX_KEY,
            "Locations",
            "tblLocations",
            f"SELECT * FROM read_csv('{csv_file}')",
        )


async def test_query_table_cannot_write_file_to_disk(ts_xlsx, tmp_path):
    out_file = tmp_path / "output.csv"
    with pytest.raises(ModelRetry):
        ts_xlsx.query_table(
            XLSX_KEY,
            "Locations",
            "tblLocations",
            f"COPY df TO '{out_file}'",
        )
    assert not out_file.exists()


async def test_query_store_and_preview_cannot_read_arbitrary_csv(ts_xlsx, tmp_path):
    csv_file = tmp_path / "secret.csv"
    csv_file.write_text("a,b\n1,2\n3,4\n")
    with pytest.raises(ModelRetry):
        ts_xlsx.query_store_and_preview(
            tables_used=[LOC_TABLE],
            query=f"SELECT * FROM read_csv('{csv_file}')",
            table_name="injected",
        )


async def test_query_store_and_preview_cannot_write_file_to_disk(ts_xlsx, tmp_path):
    out_file = tmp_path / "output.csv"
    with pytest.raises(ModelRetry):
        ts_xlsx.query_store_and_preview(
            tables_used=[LOC_TABLE],
            query=f"COPY loc TO '{out_file}'",
            table_name="out",
        )
    assert not out_file.exists()
