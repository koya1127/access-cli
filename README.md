# access-cli

A command-line tool for automating Microsoft Access (.accdb/.mdb) files — read/write VBA code, manipulate forms, execute SQL, and edit saved queries — without opening the Access GUI.

## Requirements

- Windows (x86 or x64)
- Microsoft Access installed (Access.Application COM / DAO.DBEngine.120)
- [.NET 9.0 Runtime](https://dotnet.microsoft.com/download/dotnet/9.0) (framework-dependent build)

> **Why 32-bit?** ACE OLEDB / DAO.DBEngine.120 are 32-bit in-process COM servers. The tool must run as a 32-bit process (`win-x86`).

## Installation

### From release (recommended)

Download the latest release and run `install.ps1`:

```powershell
.\install.ps1
```

This publishes to `%LOCALAPPDATA%\access-cli\` and adds it to your PATH.

### Build from source

```powershell
dotnet publish src/AccessCli/AccessCli.csproj -c Release -r win-x86 --no-self-contained -o publish_x86
```

## Commands

All commands take `<db_path>` as the first argument (path to the `.accdb` or `.mdb` file).

### VBA

| Command | Description |
|---------|-------------|
| `list-modules <db>` | List all VBA modules with type and line count |
| `read-vba <db> <module> [-o file]` | Print (or save) VBA source of a module |
| `write-vba <db> <module> <code_file>` | Overwrite a module's VBA source from a file |

### Forms

| Command | Description |
|---------|-------------|
| `list-forms <db>` | List all forms |
| `list-controls <db> <form>` | List controls in a form (name, type, caption) |
| `export-form <db> <form> <output>` | Export form definition via `SaveAsText` |
| `import-form <db> <form> <input>` | Import form definition via `LoadFromText` |
| `export-all <db> <output_dir>` | Export all modules (`.bas`) and forms (`.form`) to a directory |
| `import-all <db> <input_dir>` | Import all modules and forms from a directory |

### Tables & SQL

| Command | Description |
|---------|-------------|
| `list-tables <db>` | List all tables (including linked tables) |
| `query-sql <db> <sql> [-f file]` | Run a SELECT and print results as TSV |
| `exec-sql <db> <sql> [-f file]` | Run INSERT / UPDATE / DELETE |

Use `-f <file>` to read the SQL from a file — required when table/column names contain non-ASCII characters (see [Known Limitations](#known-limitations)).

### Saved Queries

| Command | Description |
|---------|-------------|
| `list-queries <db>` | List all saved queries with their SQL |
| `export-queries <db> <output_dir>` | Export each query's SQL to a `.sql` file in a directory |
| `get-query-sql <db> <query_name> [-o file]` | Print (or save) the SQL of a named query. Partial name match is supported. |
| `set-query-sql <db> <query_name> -f <sql_file>` | Overwrite a query's SQL from a file |

### Other

| Command | Description |
|---------|-------------|
| `set-caption <db> <old> <new>` | Binary-patch a Caption string directly in the accdb file (safe only when byte length is equal) |

## Examples

```powershell
# List tables
access-cli list-tables mydb.accdb

# Run a SELECT (ASCII table name)
access-cli query-sql mydb.accdb "SELECT * FROM Orders"

# Run a SELECT with a Japanese table name (use --sql-file)
"SELECT * FROM [受注データ]" | Out-File -Encoding UTF8 query.sql
access-cli query-sql mydb.accdb "" -f query.sql

# Export all queries to inspect them
access-cli export-queries mydb.accdb .\queries\

# Edit a saved query and write it back
access-cli get-query-sql mydb.accdb "MyQuery" -o myquery.sql
# ... edit myquery.sql ...
access-cli set-query-sql mydb.accdb "MyQuery" -f myquery.sql

# Backup all VBA and forms
access-cli export-all mydb.accdb .\backup\
```

## Known Limitations

### Japanese (non-ASCII) strings in COM arguments

Passing Japanese strings as arguments to DAO COM dispatch fails silently or with misleading errors:

```
# This will fail if the table name is Japanese:
access-cli query-sql mydb.accdb "SELECT * FROM [日本語テーブル]"
```

**Workaround:** Write the SQL to a UTF-8 file and use `-f`:

```powershell
"SELECT * FROM [日本語テーブル]" | Out-File -Encoding UTF8 q.sql
access-cli query-sql mydb.accdb "" -f q.sql
```

Internally, the tool works around this by iterating DAO collections by index and comparing names on the .NET side, so reading works correctly even for Japanese-named tables and queries.

## Architecture

- **Language:** C# / .NET 9.0 (`net9.0-windows`)
- **Target:** `win-x86` (32-bit required for DAO)
- **COM automation:** `Access.Application` (VBA / forms), `DAO.DBEngine.120` (tables / SQL / queries)
- **CLI framework:** [System.CommandLine](https://github.com/dotnet/command-line-api)

## Changelog

### v0.3.0
- Add `export-queries`: export all saved query SQLs to `.sql` files
- Add `get-query-sql`: read a query's SQL to stdout or file (supports partial name match)
- Add `set-query-sql`: overwrite a query's SQL from a file

### v0.2.0
- Add `query-sql`, `exec-sql` (INSERT via DAO AddNew/Update to handle Japanese table names)
- Add `list-tables`, `list-queries`
- Add `export-all`, `import-all`

### v0.1.0
- Initial release: `list-modules`, `read-vba`, `write-vba`, `list-forms`, `list-controls`, `export-form`, `import-form`, `set-caption`

## License

MIT
