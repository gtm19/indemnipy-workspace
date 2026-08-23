from pathlib import Path

import pytest
from indemnipy_ai.capabilities.excel._capability_state import ExcelRuntimeState
from indemnipy_ai.capabilities.excel._functions import DateParsingOptions
from indemnipy_ai.capabilities.excel._toolset import _ExcelToolset, _QueryTable
from inline_snapshot import snapshot
from pydantic_ai.exceptions import ModelRetry

TEST_DATA_DIR = Path(__file__).parent.parent.parent / "data"
XLSX_KEY = "Vantris_Pharmaceuticals_STP_Submission_2026.xlsx"
XLSX = TEST_DATA_DIR / XLSX_KEY
XLSM_KEY = "Vantris_Pharmaceuticals_STP_Submission_2026.xlsm"
XLSM = TEST_DATA_DIR / XLSM_KEY


def _make_toolset(path: Path | None = None) -> _ExcelToolset:
    state = ExcelRuntimeState()
    ts = _ExcelToolset(
        id="test",
        runtime_state=state,
        date_parsing_options=DateParsingOptions(
            parse_dates=True, relaxed_about_day=True
        ),
    )
    if path is not None:
        ts.load_workbook(path)
    return ts


# ---------------------------------------------------------------------------
# load_workbook
# ---------------------------------------------------------------------------


def test_load_workbook_returns_filename_as_key():
    ts = _make_toolset()
    key = ts.load_workbook(XLSX)
    assert key == snapshot("Vantris_Pharmaceuticals_STP_Submission_2026.xlsx")
    assert XLSX_KEY in ts.runtime_state.workbooks


def test_load_workbook_deduplicates_keys_with_numeric_suffix():
    ts = _make_toolset()
    k1 = ts.load_workbook(XLSX)
    k2 = ts.load_workbook(XLSX)
    k3 = ts.load_workbook(XLSX)
    assert k1 == snapshot("Vantris_Pharmaceuticals_STP_Submission_2026.xlsx")
    assert k2 == snapshot("Vantris_Pharmaceuticals_STP_Submission_2026.xlsx (1)")
    assert k3 == snapshot("Vantris_Pharmaceuticals_STP_Submission_2026.xlsx (2)")


# ---------------------------------------------------------------------------
# list_workbooks
# ---------------------------------------------------------------------------


def test_list_workbooks_empty():
    ts = _make_toolset()
    assert ts.list_workbooks() == snapshot("No workbooks loaded.")


def test_list_workbooks_after_load_contains_filename():
    ts = _make_toolset(XLSX)
    result = ts.list_workbooks()
    assert XLSX_KEY in result
    assert "Workbook:" in result


# ---------------------------------------------------------------------------
# list_worksheets
# ---------------------------------------------------------------------------


def test_list_worksheets_with_workbook_name():
    ts = _make_toolset(XLSX)
    result = ts.list_worksheets(XLSX_KEY)
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


def test_list_worksheets_without_name_returns_all_workbooks():
    ts = _make_toolset(XLSX)
    result = ts.list_worksheets()
    assert result is not None
    assert XLSX_KEY in result
    assert len(result[XLSX_KEY]) == 4  # Information, Claims, Locations, Sendings


def test_list_worksheets_returns_none_for_unknown_workbook():
    ts = _make_toolset(XLSX)
    assert ts.list_worksheets("does-not-exist.xlsx") is None


# ---------------------------------------------------------------------------
# list_tables_and_metadata
# ---------------------------------------------------------------------------


def test_list_tables_and_metadata():
    ts = _make_toolset(XLSX)
    result = ts.list_tables_and_metadata()
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


def test_list_tables_and_metadata_with_workbook_name_scopes_to_that_workbook():
    ts = _make_toolset(XLSX)
    result = ts.list_tables_and_metadata(XLSX_KEY)
    assert set(result.keys()) == {XLSX_KEY}
    assert "Locations" in result[XLSX_KEY]
    assert "Sendings" in result[XLSX_KEY]


def test_list_tables_and_metadata_excludes_sheets_with_no_tables():
    ts = _make_toolset(XLSX)
    sheets = ts.list_tables_and_metadata()[XLSX_KEY]
    assert "Information" not in sheets
    assert "Claims" not in sheets


def test_list_tables_and_metadata_returns_empty_dict_for_unknown_workbook():
    ts = _make_toolset(XLSX)
    assert ts.list_tables_and_metadata("no-such.xlsx") == {}


# ---------------------------------------------------------------------------
# preview_table
# ---------------------------------------------------------------------------


def test_preview_table_returns_string_containing_data():
    ts = _make_toolset(XLSX)
    result = ts.preview_table(XLSX_KEY, "Locations", "tblLocations")
    assert isinstance(result, str)
    assert "LOC-001" in result


def test_preview_table_returns_none_for_missing_workbook():
    ts = _make_toolset(XLSX)
    assert ts.preview_table("no.xlsx", "Locations", "tblLocations") is None


def test_preview_table_returns_none_for_missing_sheet():
    ts = _make_toolset(XLSX)
    assert ts.preview_table(XLSX_KEY, "NoSheet", "tblLocations") is None


def test_preview_table_returns_none_for_missing_table():
    ts = _make_toolset(XLSX)
    assert ts.preview_table(XLSX_KEY, "Locations", "noTable") is None


def test_preview_table_respects_n_rows():
    ts = _make_toolset(XLSX)
    # tblLocations has 32 data rows; request only 2
    result = ts.preview_table(XLSX_KEY, "Locations", "tblLocations", n_rows=2)
    assert result is not None
    assert result.count("LOC-") == 2


def test_preview_table_respects_offset():
    ts = _make_toolset(XLSX)
    result_from_start = ts.preview_table(
        XLSX_KEY, "Locations", "tblLocations", n_rows=1, offset=0
    )
    result_offset = ts.preview_table(
        XLSX_KEY, "Locations", "tblLocations", n_rows=1, offset=1
    )
    assert result_from_start is not None
    assert result_offset is not None
    assert result_from_start != result_offset


# ---------------------------------------------------------------------------
# get_range
# ---------------------------------------------------------------------------


def test_get_range_returns_list_of_lists():
    ts = _make_toolset(XLSX)
    result = ts.get_range(XLSX_KEY, "Information", "B1:C3")
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


def test_get_range_returns_none_for_missing_workbook():
    ts = _make_toolset(XLSX)
    assert ts.get_range("no.xlsx", "Information", "B1:C3") is None


def test_get_range_raises_model_retry_for_invalid_range():
    ts = _make_toolset(XLSX)
    with pytest.raises(ModelRetry):
        ts.get_range(XLSX_KEY, "Information", "NOT_A_RANGE")


# ---------------------------------------------------------------------------
# add_table_from_range
# ---------------------------------------------------------------------------


def test_add_table_from_range_success():
    ts = _make_toolset(XLSX)
    result = ts.add_table_from_range(XLSX_KEY, "Locations", "A1:C3", "myTable")
    assert result == snapshot(
        {
            "myTable": {
                "Location ID": "String",
                "Site Name": "String",
                "Address": "String",
            }
        }
    )
    # Table is now visible to other toolset methods
    assert ts.preview_table(XLSX_KEY, "Locations", "myTable") is not None


def test_add_table_from_range_returns_none_for_missing_workbook():
    ts = _make_toolset(XLSX)
    assert ts.add_table_from_range("no.xlsx", "Locations", "A1:C3", "t") is None


def test_add_table_from_range_raises_model_retry_for_duplicate_table_name():
    ts = _make_toolset(XLSX)
    with pytest.raises(ModelRetry, match="already exists"):
        ts.add_table_from_range(XLSX_KEY, "Locations", "A1:C3", "tblLocations")


# ---------------------------------------------------------------------------
# query_table
# ---------------------------------------------------------------------------


def test_query_table_returns_formatted_string():
    ts = _make_toolset(XLSX)
    result = ts.query_table(
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


def test_query_table_returns_none_for_missing_workbook():
    ts = _make_toolset(XLSX)
    assert ts.query_table("no.xlsx", "Locations", "tblLocations", "SELECT 1") is None


def test_query_table_returns_none_for_missing_sheet():
    ts = _make_toolset(XLSX)
    assert ts.query_table(XLSX_KEY, "NoSheet", "tblLocations", "SELECT 1") is None


def test_query_table_returns_none_for_missing_table():
    ts = _make_toolset(XLSX)
    assert ts.query_table(XLSX_KEY, "Locations", "noTable", "SELECT 1") is None


def test_query_table_raises_model_retry_for_bad_sql():
    ts = _make_toolset(XLSX)
    with pytest.raises(ModelRetry):
        ts.query_table(
            XLSX_KEY, "Locations", "tblLocations", "SELECT FROM WHERE BROKEN"
        )


# ---------------------------------------------------------------------------
# get_workbook_vba
# ---------------------------------------------------------------------------


def test_get_workbook_vba_returns_markdown_for_xlsm():
    ts = _make_toolset(XLSM)
    result = ts.get_workbook_vba(XLSM_KEY)
    assert isinstance(result, str)
    assert result.startswith(
        "# VBA Macros in Vantris_Pharmaceuticals_STP_Submission_2026.xlsm"
    )
    assert "## Analysis Results" in result
    assert "```vba" in result


def test_get_workbook_vba_returns_none_for_xlsx_without_macros():
    ts = _make_toolset(XLSX)
    assert ts.get_workbook_vba(XLSX_KEY) is None


def test_get_workbook_vba_returns_none_for_unknown_workbook():
    ts = _make_toolset(XLSX)
    assert ts.get_workbook_vba("no.xlsm") is None


# ---------------------------------------------------------------------------
# query_store_and_preview
# ---------------------------------------------------------------------------


def test_query_store_and_preview_stores_result_and_returns_query_result():
    ts = _make_toolset(XLSX)
    result = ts.query_store_and_preview(
        tables_used=[
            _QueryTable(
                table="tblLocations",
                reference_name="loc",
                workbook=XLSX_KEY,
                sheet="Locations",
            )
        ],
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
    assert "uk_locs" in ts.runtime_state.derived_tables


def test_query_store_and_preview_raises_model_retry_when_sheet_missing_for_workbook_table():
    ts = _make_toolset(XLSX)
    with pytest.raises(ModelRetry, match="sheet is required"):
        ts.query_store_and_preview(
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


def test_query_store_and_preview_raises_model_retry_for_missing_workbook_table():
    ts = _make_toolset(XLSX)
    with pytest.raises(ModelRetry, match="not found"):
        ts.query_store_and_preview(
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


def test_query_store_and_preview_raises_model_retry_for_missing_derived_table():
    ts = _make_toolset(XLSX)
    with pytest.raises(ModelRetry, match="not found"):
        ts.query_store_and_preview(
            tables_used=[_QueryTable(table="ghost", reference_name="g")],
            query="SELECT * FROM g",
            table_name="out",
        )


def test_query_store_and_preview_raises_model_retry_for_bad_sql():
    ts = _make_toolset(XLSX)
    with pytest.raises(ModelRetry):
        ts.query_store_and_preview(
            tables_used=[
                _QueryTable(
                    table="tblLocations",
                    reference_name="t",
                    workbook=XLSX_KEY,
                    sheet="Locations",
                )
            ],
            query="SELECT FROM WHERE BROKEN",
            table_name="out",
        )


# ---------------------------------------------------------------------------
# list_derived_tables
# ---------------------------------------------------------------------------


def test_list_derived_tables_is_empty_before_any_queries():
    ts = _make_toolset(XLSX)
    assert ts.list_derived_tables() == {}


def test_list_derived_tables_populated_after_query_store_and_preview():
    ts = _make_toolset(XLSX)
    ts.query_store_and_preview(
        tables_used=[
            _QueryTable(
                table="tblLocations",
                reference_name="t",
                workbook=XLSX_KEY,
                sheet="Locations",
            )
        ],
        query='SELECT "Location ID", "Country" FROM t LIMIT 5',
        table_name="sample",
    )
    assert ts.list_derived_tables() == snapshot(
        {"sample": {"Location ID": "String", "Country": "String"}}
    )


# ---------------------------------------------------------------------------
# preview_derived_table
# ---------------------------------------------------------------------------


def test_preview_derived_table_returns_none_for_unknown_table():
    ts = _make_toolset(XLSX)
    assert ts.preview_derived_table("ghost") is None


def test_preview_derived_table_returns_string_for_known_table():
    ts = _make_toolset(XLSX)
    ts.query_store_and_preview(
        tables_used=[
            _QueryTable(
                table="tblLocations",
                reference_name="loc",
                workbook=XLSX_KEY,
                sheet="Locations",
            )
        ],
        query='SELECT "Location ID", "Country" FROM loc WHERE "Country" = \'United Kingdom\'',
        table_name="uk_locs",
    )
    result = ts.preview_derived_table("uk_locs", n_rows=2)
    assert result == snapshot(
        " Location ID  Country        \n"
        " ---          ---            \n"
        " str          str            \n"
        " LOC-001      United Kingdom \n"
        " LOC-002      United Kingdom "
    )


def test_preview_derived_table_respects_n_rows_and_offset():
    ts = _make_toolset(XLSX)
    ts.query_store_and_preview(
        tables_used=[
            _QueryTable(
                table="tblLocations",
                reference_name="loc",
                workbook=XLSX_KEY,
                sheet="Locations",
            )
        ],
        query='SELECT "Location ID", "Country" FROM loc WHERE "Country" = \'United Kingdom\'',
        table_name="uk_locs",
    )
    r1 = ts.preview_derived_table("uk_locs", n_rows=1, offset=0)
    r2 = ts.preview_derived_table("uk_locs", n_rows=1, offset=1)
    assert r1 is not None and r2 is not None
    assert r1 != r2
