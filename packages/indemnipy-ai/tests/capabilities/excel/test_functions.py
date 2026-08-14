from datetime import datetime, timezone
from typing import Any

import pytest
from indemnipy_ai.capabilities.excel._functions import (
    _gently_parse_datetime as gently_parse_datetime,
)


@pytest.mark.parametrize(
    "relaxed_about_day",
    [True, False],
    ids=["relaxed_about_day=True", "relaxed_about_day=False"],
)
@pytest.mark.parametrize(
    "input_value, expected_output",
    [
        ("2023-01-15", datetime(2023, 1, 15, tzinfo=timezone.utc)),
        ("15/01/2023", datetime(2023, 1, 15, tzinfo=timezone.utc)),
        ("January 15, 2023", datetime(2023, 1, 15, tzinfo=timezone.utc)),
        ("15 Jan 2023", datetime(2023, 1, 15, tzinfo=timezone.utc)),
        ("01-15-2023", datetime(2023, 1, 15, tzinfo=timezone.utc)),
        ("2023/01/15 4:10:12", datetime(2023, 1, 15, 4, 10, 12, tzinfo=timezone.utc)),
        ("2023/01/15 4:30am", datetime(2023, 1, 15, 4, 30, tzinfo=timezone.utc)),
        ("2023.01.15 4am", datetime(2023, 1, 15, 4, tzinfo=timezone.utc)),
    ],
)
def test_should_always_parse(
    input_value: Any, expected_output: datetime, relaxed_about_day: bool
):
    """
    Tests for inputs which should always parsed as datetimes regardless of the relaxed_about_day setting.
    """
    result = gently_parse_datetime(input_value, relaxed_about_day=relaxed_about_day)
    assert result == expected_output


@pytest.mark.parametrize(
    "relaxed_about_day",
    [True, False],
    ids=["relaxed_about_day=True", "relaxed_about_day=False"],
)
@pytest.mark.parametrize(
    "input_value",
    [
        "2/4/2023",  # Ambiguous date, should return original string
        "2023-13-01",  # Invalid month, should return original string
        "2023-02-30",  # Invalid day, should return original string
        "Back in 2021",  # No month or day, should return original string
        "Back in year of yonder yore",  # No month or day, should return original string
        20250125,  # Not a valid date string, should return original integer
        None,  # None value, should return None
        "",  # Empty string, should return empty string
    ],
)
def test_should_never_parse(input_value: Any, relaxed_about_day: bool):
    """
    Tests for inputs which should never be parsed as datetimes.
    """
    result = gently_parse_datetime(input_value, relaxed_about_day=relaxed_about_day)
    assert result == input_value


@pytest.mark.parametrize(
    "input_value, expected_output",
    [
        ("October 2025", datetime(2025, 10, 1, tzinfo=timezone.utc)),
        ("2020 Apr", datetime(2020, 4, 1, tzinfo=timezone.utc)),
    ],
)
def test_should_parse_only_when_relaxed(input_value: Any, expected_output: datetime):
    """
    Tests for inputs which should only be parsed as datetimes when relaxed_about_day is True.
    """
    # Test with relaxed_about_day=True (must parse)
    result_relaxed = gently_parse_datetime(input_value, relaxed_about_day=True)
    assert result_relaxed == expected_output

    # Test with relaxed_about_day=False (must not parse)
    result_strict = gently_parse_datetime(input_value, relaxed_about_day=False)
    assert result_strict == input_value
