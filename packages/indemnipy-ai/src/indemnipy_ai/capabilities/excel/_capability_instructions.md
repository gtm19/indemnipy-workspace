## Working with Spreadsheets

This Capability makes it easy for you to work with spreadsheets.

Specifically, you will be able to:
- Load spreadsheets and inspect their contents
- Query tables (which are automatically added when there are named tables already in the workbook) using SQL (duckdb) syntax
- Inspect the content of worksheets
- Add new tables to your internal storage based on a workbook, worksheet, range and name
- Store query results as derived tables — summaries or joins across workbooks and sheets — without needing to tie them to any specific workbook or worksheet

Follow this workflow in order. Do not skip steps.

### Step 1 - View Loaded Workbooks / Load Some Workbooks

It is possible that some workbooks have already been loaded into the session. You can view all loaded workbooks by calling `list_workbooks()`. This will return a list of workbook names.

If the data returned is sufficient for your task, you can skip to Step 2. Otherwise, you will need to load additional workbooks.

You should **never** attempt to load a workbook that has already been loaded.

If the user makes reference to a workbook in their input (i.e. a file with an `.xls`, `.xlsx`, or `.xlsm` extension),
you should first load the workbook using `load_workbook()`.

This will make the workbook available for further inspection and querying.

### Step 2 - Discover

Data which might be useful to your task might already have been loaded in if it was
a properly formatted and named table in the workbook. You can discover all available tables across loaded workbooks by calling `list_tables_and_metadata()`.

This returns a nested structure: `{workbook_name: {sheet_name: {table_name: {col: (dtype,)}}}}`.
Only sheets that contain at least one table are included.

There may, however, be "table-like" structures in the workbook which are not named tables. You can inspect the contents of worksheets by calling `list_worksheets()` and then `get_range()` on each relevant worksheet. This will allow you to see the contents of the worksheet and identify any table-like structures which you may want to add as a named table for querying. Listing worksheets will also include the range of the worksheet, which should help with identifying the outer limits of where to look for table-like structures.

** IMPORTANT **: Be conservative when adding new tables from ranges. If tables are already loaded from a sheet, prefer using those rather than creating overlapping range-based tables for the same data.

### Step 3 - Inspect / Validate

Once we have all the tables we need, we can inspect them using `preview_table(workbook_name, sheet_name, table_name)`. This will allow us to see the column names and raw formats (string, number, date, etc.) of each table. We can also spot columns that will need type coercion before the final query.

**Workbook tables are read-only.** Any transformation, cleaning, or conforming must be done via `query()` — write the result to a derived table, then work from that. Do not attempt to modify a workbook table in place.

### Conforming date columns

This is especially important for date columns. For every column that contains dates, check for unparseable values with a single query **before** doing anything else with that column:

```sql
SELECT * FROM df WHERE TRY_CAST("<column>" AS DATE) IS NULL
```

If this returns no rows, all dates are valid. If rows are returned, inspect them to understand the format and resolve via an appropriate `COALESCE` / `TRY_STRPTIME` in your `query()` call — either coerce with the correct format string, or flag the row — before proceeding. Store the cleaned result as a derived table and use that for all subsequent work.

### Final Step

Once all tables are loaded and inspected:

- Use `query()` to transform, clean, join, or summarise — across any combination of workbook tables and derived tables. Results are stored as a **derived table** in the session store and are not tied to any workbook or worksheet.
- If you only need to read a single table without storing the result, use `query_table(workbook_name, sheet_name, table_name, query)`.

#### Working with derived tables

Derived tables persist for the duration of the session and can be:
- Listed with `list_derived_tables()` — returns `{table_name: {col: (dtype,)}}`
- Inspected with `preview_derived_table(table_name)`
- Used as inputs to further `query()` calls by omitting `workbook` and `sheet` from the `QueryTable` entry

When composing a `query()` call, each `QueryTable` entry specifies:
- `table`: the table name
- `reference_name`: the SQL alias used in the query
- `workbook` (optional): the workbook name — omit when referencing a derived table
- `sheet` (optional): the sheet name — required when `workbook` is specified

You should always try to do the most efficient query to answer the question, given that returning large amounts of data can be slow and expensive.


### Rules
- Always validate date columns (Step 3) before using them in output.
- Never attempt to enumerate or reconstruct row data from memory — always query for it.
- Prefer explicit `CAST` over relying on implicit string parsing.
- When calling `preview_table` or `query_table`, always supply `workbook_name`, `sheet_name`, **and** `table_name` — all three are required.
- When looking up using `workbook_name`, it is the _name_ of the workbook (i.e. just `myworkbook.xlsx`), not the full path.
- Never try to overwrite a workbook table. All transformations must go through `query()` into derived tables.
- `add_table_from_range` requires a name that does not already exist in that sheet.
