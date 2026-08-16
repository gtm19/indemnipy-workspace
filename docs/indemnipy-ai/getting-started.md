---
icon: lucide/rocket
---

# Getting Started

## Installation

```bash
pip install indemnipy-ai
```

Or with [uv](https://docs.astral.sh/uv/):

```bash
uv add indemnipy-ai
```

## Basic usage

The minimal setup is an `Agent` with `ExcelCapability` attached. No other configuration is required.

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

Without any spreadsheets pre-loaded, the agent will report that none are available. You can ask it to load a file by path, or you can pre-load files yourself — see below.

---

## Pre-loading Spreadsheets

Pass local file paths to `ExcelRuntimeState` before the run. The capability will load the spreadsheets immediately and the agent will have access to them from the first message.

```python
from dataclasses import dataclass, field
from pathlib import Path

from indemnipy_ai.capabilities.excel import ExcelCapability, ExcelRuntimeState
from pydantic_ai import Agent


@dataclass
class Deps:
    excel_runtime_state: ExcelRuntimeState = field(default_factory=ExcelRuntimeState)


SPREADSHEETS = [
    Path("data/STP_Submission_2026.xlsx"),
]

agent = Agent(
    "openai:gpt-4o",
    capabilities=[ExcelCapability()],
    output_type=str,
    deps_type=Deps,
)


def main():
    deps = Deps(excel_runtime_state=ExcelRuntimeState(spreadsheets=SPREADSHEETS))
    result = agent.run_sync(
        "Extract the claims data by year — total count and total value.",
        deps=deps,
    )
    print(result.output)


if __name__ == "__main__":
    main()
```

!!! note "Files must be local"
    Spreadsheets must exist on the local filesystem. Network paths and cloud storage URLs are not supported.

You can also omit `spreadsheets` entirely and let the agent load workbooks on demand — it will call `load_workbook` itself when you refer to a file by path in your message.

---

## Passing deps (recommended)

Passing a `Deps` object is **recommended** if you want to:

- Access derived tables or workbooks after the run
- Preserve state across turns in a multi-turn conversation

The `Deps` class just needs to have an `excel_runtime_state` attribute of type `ExcelRuntimeState`. A dataclass works well:

```python
@dataclass
class Deps:
    excel_runtime_state: ExcelRuntimeState = field(default_factory=ExcelRuntimeState)
```

After `agent.run_sync(...)` returns, you can inspect everything the agent created:

```python
deps = Deps(excel_runtime_state=ExcelRuntimeState(spreadsheets=SPREADSHEETS))
result = agent.run_sync("Summarise claims by year.", deps=deps)

# Derived tables created during the run
for name, table in deps.excel_runtime_state.derived_tables.items():
    print(f"Derived table: {name}")
    print(table.dataframe)

# The raw workbooks that were loaded
for name, workbook in deps.excel_runtime_state.workbooks.items():
    print(f"Workbook: {name}")
    for sheet in workbook.sheets:
        print(f"  Sheet: {sheet.name}")
        for table in sheet.tables:
            print(f"    Table: {table.name}")
            print(table.dataframe)
```

If you do not pass `deps`, the capability creates a fresh `ExcelRuntimeState` internally. The agent still works, but you cannot access any derived tables or loaded workbooks after the run, and state is lost between turns.

---

## Multi-turn conversations

Pass the same `deps` instance to each `agent.run()` call to preserve workbooks and derived tables across turns:

```python
deps = Deps(excel_runtime_state=ExcelRuntimeState(spreadsheets=SPREADSHEETS))

result1 = agent.run_sync("Load the workbook and list the available tables.", deps=deps)
print(result1.output)

result2 = agent.run_sync("Now summarise claims by year.", deps=deps)
print(result2.output)
```

The agent will remember which workbooks are loaded and which derived tables exist from the previous turn.

---

## Running as a CLI

`pydantic-ai` agents support an interactive CLI mode. Call `agent.to_cli_sync()` instead of `agent.run_sync()` to drop into a REPL-style loop:

```python
def cli():
    deps = Deps(excel_runtime_state=ExcelRuntimeState(spreadsheets=SPREADSHEETS))
    agent.to_cli_sync(deps=deps)


if __name__ == "__main__":
    cli()
```

You can also run this directly from the command line if you structure your script with a `cli()` entry point and call it from `__main__`.

---

## Observability with Logfire

[Logfire](https://logfire.pydantic.dev/) integrates directly with pydantic-ai and gives you traces for every agent run, tool call, and model request. To enable it, add the following before your agent setup:

```python
import logfire

logfire.configure()
logfire.instrument_pydantic_ai()
```

For system-level metrics (CPU, memory, etc.):

```python
logfire.instrument_system_metrics()
```

Logfire is a dev dependency of indemnipy-ai and is not required at runtime.
