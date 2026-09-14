# Excel VBA SQL Emulator Class

A lightweight, in-memory SQL-like query engine for standard Excel VBA.

The project works with Excel `Range` objects and 2D `Variant` arrays. Queries are built with a fluent VBA API; the library does **not** parse SQL text. SQL snippets in the documentation are conceptual equivalents of the VBA operations.

## Features

- Load tabular data from a `Range` or a 2D `Variant` array.
- Select columns by name or 1-based index.
- Use aliases with `AS`.
- Filter with `WHERE`, `AND`, `LIKE`, `IN`, `IS NULL`, and `IS NOT NULL`.
- Use typed conversion: `TEXT`, `LONG`, `DOUBLE`, `CURRENCY`, `DATE`, and `BOOLEAN`.
- Sort by one or more columns with `ASC` or `DESC`.
- Remove duplicates with `DISTINCT`.
- Use `LIMIT`, `OFFSET`, and `Page`.
- Perform `INNER JOIN` and `LEFT JOIN` operations.
- Group rows and calculate `COUNT`, `SUM`, `AVG`, `MIN`, and `MAX`.
- Return results as 2D `Variant` arrays or write them to an Excel range.
- Access a table directly through the `SqlTable` facade.

## Requirements

- Microsoft Excel with VBA support.
- No third-party libraries, DLLs, ActiveX controls, or mandatory external references.
- `Scripting.Dictionary` is used through late binding.

## Installation

1. Open or create an Excel workbook and save it as `.xlsm` if it will contain macros.
2. Open the VBA editor with `Alt+F11`.
3. Import `SqlQuery.cls` as a **Class Module**.
4. Import `SqlTable.cls` if direct table access is required.
5. Import `SqlMutation.cls` if `INSERT`/`UPDATE` operations are required.
6. Select **Debug → Compile VBAProject**.

The three `.cls` files are runtime **Class Modules**, not Standard Modules. Import them with **File → Import File** in the VBA editor; do not paste their contents into a standard module.

The `.cls` files use Windows CRLF line endings and contain the VBA class metadata required by the VBA importer. The repository includes `.gitattributes` with `*.cls -text` so Git preserves those line endings. If a downloaded file has been rewritten by another tool, restore CRLF before importing it.

## SqlMutation: INSERT and UPDATE

`SqlMutation` is a separate write facade. It changes an in-memory copy first and writes to Excel only after `Commit`.

```vba
Dim mutation As SqlMutation
Dim values As Object

Set mutation = New SqlMutation
Set mutation = mutation.AutoRange("Employees")

Set values = CreateObject("Scripting.Dictionary")
values("EmployeeId") = 13
values("Name") = "Nikolay"
values("Department") = "IT"
values("Salary") = 3100
values("Active") = True

mutation.Insert values
Debug.Print mutation.RowsAffected       ' 1
mutation.Commit
```

For an update:

```vba
Set mutation = New SqlMutation
Set mutation = mutation.AutoRange("Employees")

Set values = CreateObject("Scripting.Dictionary")
values("Salary") = 3600

mutation.UpdateWhere "Name", "=", "Ivan", values
Debug.Print mutation.RowsAffected
mutation.Commit
```

`OpenRange` remains available for an explicitly controlled range. `Rollback` reloads the original worksheet range and discards uncommitted in-memory changes. `DeleteWhere` removes matching rows from the in-memory snapshot; the worksheet changes only after `Commit`. After a DELETE commit, trailing worksheet rows from the old target range are physically deleted so that stale rows are not left below the compacted result. `SqlMutation` exposes `HasPendingChanges`: it is True after INSERT/UPDATE/DELETE and False after Commit/Rollback. Commit keeps a backup of the original range values and attempts to restore them if the write fails. Commit history is available via `GetCommitLog` and can be cleared with `ClearCommitLog`. The log columns are `Timestamp`, `Worksheet`, `Range`, `RowDelta`, `Status` and `Error`. For multiple conditions use the mutation builder: `mutation.Where("Department", "=", "IT").AndWhere "Salary", ">", 3000`, then inspect `mutation.MatchedCount` and call `mutation.Update values` or `mutation.Delete`. Update/Delete require at least one filter; use `Truncate True` for all data rows. The builder also supports `OrWhere`, `NotWhere`, `BeginGroup`, `EndGroup`, `Count` and `Preview`.

### SqlMutation diagnostics`r`n`r`n```vba`rnmutation.SetConversionPolicy "IGNORE"`r`nmutation.EnableDiagnostics`r`nmutation.WhereTyped "Salary", ">", "3000", "DOUBLE"`r`nDebug.Print mutation.MatchedCount`r`nerrors = mutation.GetConversionErrors()`r`n````r`n`r`n`GetConversionErrors` връща `SourceRow`, `Context`, `ValueType`, `Policy` и `Value`. Изчиства се с `ClearDiagnostics`.`r`n`r`n`SqlTable.FindFirst` returns a late-bound row dictionary. Print a field from the dictionary, not the object itself:

```vba
Dim row As Object
Set row = employees.FindFirst("Name", "=", "Ivan")
If row Is Nothing Then
    Debug.Print "Ivan was not found."
Else
    Debug.Print row("Name"), row("Department"), row("Salary")
End If
```

## SqlTable: manual or automatic worksheet range

The manual option remains available:

```vba
Dim employees As SqlTable
Set employees = New SqlTable
Set employees = employees.OpenRange( _
    ThisWorkbook.Worksheets("Employees").Range("A1:E13"))
```

For a universal table whose data starts at `A1`, use `AutoRange` with only the worksheet name:

```vba
Set employees = New SqlTable
Set employees = employees.AutoRange("Employees")
```

`AutoRange` searches the current workbook worksheet for the last non-empty/formula cell by row and by column, then opens `A1:lastRow:lastColumn`. It assumes that the first row contains headers and that the data area does not contain completely blank rows. The header row must start at `A1` and every detected header cell must contain a title. An empty worksheet or an unknown worksheet name raises an error. `OpenRange` is still recommended when the exact range must be controlled explicitly.

## Quick start: SqlQuery

```vba
Option Explicit

Public Sub RunDepartmentQuery()
    On Error GoTo ErrorHandler

    Dim query As SqlQuery
    Set query = New SqlQuery

    Set query = query.LoadRange( _
        ThisWorkbook.Worksheets("Employees").Range("A1:E13"))
    Set query = query.SelectColumns("Name", "Salary")
    Set query = query.Where("Department", "=", "IT")
    Set query = query.OrderBy("Salary", "DESC")

    query.WriteToRange ThisWorkbook.Worksheets("Results").Range("A1")
    Exit Sub

ErrorHandler:
    MsgBox "Query failed: " & Err.Description, vbCritical
End Sub
```

The conceptual SQL equivalent is:

```sql
SELECT Name, Salary
FROM Employees
WHERE Department = 'IT'
ORDER BY Salary DESC;
```

## Quick start: SqlTable

```vba
Option Explicit

Public Sub ReadEmployeeTable()
    On Error GoTo ErrorHandler

    Dim employees As SqlTable
    Dim row As Object

    Set employees = New SqlTable
    Set employees = employees.OpenRange( _
        ThisWorkbook.Worksheets("Employees").Range("A1:E13"))

    Set row = employees.GetRowDictionary(1)
    Debug.Print row("Name")
    Debug.Print row("Salary")
    Exit Sub

ErrorHandler:
    MsgBox "Table access failed: " & Err.Description, vbCritical
End Sub
```

`SqlTable` uses 1-based **data-row** indexes. The header row is not counted.

## Source data convention

Both classes expect the first row of the source range or array to contain column headers. A typical workbook layout is:

```text
Employees!A1:E13       Main employee table
Departments!A1:D5      Department lookup table
Results!A1             Optional output location
```

The included examples use the following columns:

```text
Employees:    EmployeeId, Name, Department, Salary, Active
Departments:  Department, Manager, Office, Budget
```

## Documentation

- [`SqlQuery_examples.html`](SqlQuery_examples.html) — complete examples for logical expressions, computed columns, aggregates, composite JOINs, chained JOINs, and the public `SqlQuery` API.
- [`SqlTable_examples.html`](SqlTable_examples.html) — complete general examples and one copy-and-paste example for every public `SqlTable` method.
- [`sqlQuery_manual.html`](sqlQuery_manual.html) — English API manual for `SqlQuery`, including the current logical, aggregate, JOIN, and computed-column extensions.
- [`SqlTable_manual.html`](SqlTable_manual.html) — English API manual for `SqlTable`.
- [`SqlMutation_manual.html`](SqlMutation_manual.html) — English API manual for safe snapshot mutations and atomic commits.
- [`SqlMutation_examples.html`](SqlMutation_examples.html) — mutation workflows, validation, rollback, and commit failure examples.

Each HTML document includes the source tables, a linked contents section, complete VBA procedures, conceptual SQL equivalents, and expected behavior.

## Important behavior

- `SqlQuery` and `SqlTable` operate in memory.
- The source header row is preserved in query results.
- Empty results still contain a header row.
- Query methods return the same `SqlQuery` object and are intended for fluent composition.
- `STRICT` is the default conversion policy for explicit typed operations.
- `IGNORE` and `NULL` provide alternatives for invalid or missing typed values.
- `LIKE` uses the VBA `Like` operator, not regular expressions.
- There is no text SQL parser.
- The engine does not parse SQL text. Logical `OR`, `NOT`, parenthesized predicates, computed columns (`UPPER`, `LOWER`, `LEN`), aggregate aliases, composite JOIN keys, and chained JOINs are available through the VBA API.

## Conceptual execution order

```text
JOIN
→ WHERE
→ GROUP BY and aggregates
→ projection and aliases
→ DISTINCT
→ ORDER BY
→ LIMIT/OFFSET
```

## Current advanced API

### Logical expressions

```vba
query.BeginGroup
query.Where "Department", "=", "IT"
query.OrWhere "Department", "=", "HR"
query.EndGroup
query.AndWhere "Active", "=", True
```

The expression precedence is `NOT`, then `AND`, then `OR`. Typed variants include `OrWhereTyped` and `NotWhereTyped`.

### Computed columns

```vba
query.SelectColumns "Name"
query.SelectComputed "UPPER(Name)", "NameUpper"
query.AddComputed "LOWER(Name)", "NameLower"
query.AddComputed "LEN(Name)", "NameLength"
```

Supported computed functions are `UPPER`, `LOWER`, and `LEN`. Computed aliases can be used by `ORDER BY`, `DISTINCT`, and paging.

### Composite and chained JOINs

```vba
query.InnerJoin lookup, _
    Array("Country", "Code"), _
    Array("Country", "Code"), _
    "TEXT", "Lookup"

query.LeftJoin departments, "Department", "Department", "TEXT", "Dept"
query.LeftJoin offices, "Dept.Code", "Code", "TEXT", "Office"
```

Duplicate right-side headers receive an alias such as `Right.Department`, or an explicit prefix such as `Dept.Department`.

### Safe mutations

`SqlMutation` works on an in-memory snapshot. `Commit` validates the target range, detects external changes, keeps a backup, restores on write failure, and records `SUCCESS`/`FAILED` commit log entries. `Rollback` discards pending changes.

## License

No license file is currently included. Add the license that matches your intended distribution terms before publishing a formal release.