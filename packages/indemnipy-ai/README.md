# Indemnipy AI

This is the Indemnipy AI package, which provides a set of tools and capabilities
using AI in the wonderful world of insurance. Which inevitably means spreadsheets.

> [!NOTE]
> This package is pre-release. More features and capabilities are planned, and the API may change at any time. Please use with caution - and let us know if you have any feedback or suggestions!

## Installation

You can install the package using pip:

```bash
pip install indemnipy-ai
```

or, with `uv`:

```bash
uv pip install indemnipy-ai
# or
uv add indemnipy-ai
```


## Capabilities

This package provides an `ExcelCapability` that allows you to build Pydantic AI agents which have the ability to work with Excel files (both `.xlsx` and `.xlsm`). You can use this capability to read, write, and analyse the data in Excel files.

This capability adds the following abilities to your agent:

- Automatically discovers named tables and reads them in as Polars dataframes
- Browses sheets for "table-like" data and can create new tables (again as Polars dataframes) from specific ranges
- Reads arbitrary ranges of sheets for non-tabular data
- Non-destructively queries tables using duckdb across sheets and workbooks to create new tables - either to use to help respond to the user or to pass back to the user at the end of a run for their own uses
- Performs automatic conversion (on data load) of strings which can be unambiguously parsed as datetimes, and optionally does this for dates with missing days (pinning them to the first of the month)
- Parses any VBA code within the workbook, along with a basic security analysis

Read more about Capabilities in the [Pydantic documentation](https://pydantic.dev/docs/ai/capabilities/overview/).

### Basic Usage

```python
from indemnipy_ai.capabilities.excel import ExcelCapability
from pydantic_ai import Agent

agent = Agent(
    "openai:gpt-4o",
    capabilities=[ExcelCapability()],
    output_type=str,
)

result = agent.run_sync("What spreadsheets do you have access to?")
print(result.output)
```
