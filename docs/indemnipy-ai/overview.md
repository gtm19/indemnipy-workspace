---
icon: lucide/sparkles
---

# indemnipy-ai

**indemnipy-ai** provides [pydantic-ai](https://ai.pydantic.dev/) Capabilities for working with insurance data.

## What are Capabilities?

Capabilities are reusable bundles of tools and system instructions that you attach to a pydantic-ai `Agent`. They extend what an agent can do without you needing to define tools manually. See the [pydantic-ai Capabilities documentation](https://ai.pydantic.dev/capabilities/) for full details.

## ExcelCapability

`ExcelCapability` gives an agent the ability to work with Excel files — `.xlsx` and `.xlsm` — using a structured, step-by-step workflow.

When this capability is active, the agent can:

- **Load workbooks** — by file path (files must be local)
- **Inspect worksheets** — browse sheet names, cell ranges, and dimension metadata
- **Discover tables** — named Excel tables are picked up automatically; unformatted ranges can be promoted to named tables on demand
- **Preview and query** — read and aggregate table data using DuckDB SQL, without modifying the source file
- **Derive and store results** — store query results as named derived tables that persist across turns in a multi-turn conversation

### Agent tools

The capability registers the following tools on the agent:

| Tool | Description |
|------|-------------|
| `load_workbook` | Load a local Excel file into the session |
| `list_workbooks` | Summarise all loaded workbooks |
| `list_worksheets` | List sheet metadata for a workbook |
| `list_tables_and_metadata` | Return table names and column schemas |
| `get_range` | Read a raw cell range as a list of lists |
| `add_table_from_range` | Register an unformatted range as a named table |
| `preview_table` | Preview the first N rows of a workbook table |
| `query_table` | Run a read-only DuckDB SQL query on a workbook table |
| `query_store_and_preview` | Run a SQL query, store the result as a derived table, and return a preview |
| `list_derived_tables` | List all derived tables created in the session |
| `preview_derived_table` | Preview rows from a derived table |
| `get_workbook_vba` | Return a Markdown summary of any detected VBA macros |

### Built-in instructions

The capability ships with its own system instructions that guide the agent through a recommended workflow: discover what is loaded, inspect available tables, validate date columns, then query and derive results. You do not need to prompt the agent to follow this workflow — it is included automatically.
