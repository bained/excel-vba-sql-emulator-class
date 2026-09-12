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
5. Select **Debug → Compile VBAProject**.

The two class modules are the runtime files. The HTML files are documentation and copy-and-paste examples.

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

- [`SqlQuery_examples.html`](SqlQuery_examples.html) — complete general examples and one copy-and-paste example for every public `SqlQuery` method.
- [`SqlTable_examples.html`](SqlTable_examples.html) — complete general examples and one copy-and-paste example for every public `SqlTable` method.
- [`sqlQuery_manual.html`](sqlQuery_manual.html) — English API manual for `SqlQuery`.
- [`SqlTable_manual.html`](SqlTable_manual.html) — English API manual for `SqlTable`.

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
- There is no support yet for SQL text, `OR`, `NOT`, parenthesized predicates, computed columns, or functions such as `UPPER`, `LOWER`, and `LEN`.

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

## License

No license file is currently included. Add the license that matches your intended distribution terms before publishing a formal release.