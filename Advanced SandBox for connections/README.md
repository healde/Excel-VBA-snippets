# Advanced Sandbox for Excel connections

> A VBA sandbox for exploring how Excel imports data from another workbook through Power Query/M formulas, workbook connections, the Data Model, and OLE DB providers.

## TL;DR

This project is a learning workbook, not a production add-in. Use it to see what Excel creates behind the scenes when one workbook imports data from another.

Typical flow:

1. Open `SandBox1.xlsm` in Excel and enable macros only in a trusted test environment.
2. Run `ChangeSettings` to set the external workbook path, target sheet/table, provider, query name, and connection name.
3. Run `ExportAsDistant` to generate a sample external workbook with configured sheets and column headers.
4. Run `DoQueries` to create or update a Power Query/M query.
5. Run either `AddConnectionWithDataModel` or `AddConnectionOnly` to create a workbook connection.
6. Run `ImportFromConnection` or `ImportFromExternal` to load the external data into the active worksheet.
7. Use `showConnections`, `showListobjects`, `cleanConnections`, and `cleanListObjects` to inspect or reset the workbook state.

The sandbox compares two main provider routes:

- `Microsoft.Mashup.OleDb.1` — Power Query/M-based access through workbook queries.
- `Microsoft.ACE.OLEDB.12.0` — direct OLE DB access to workbook tables/ranges.

---

## For readers who are new to Excel

Excel is a spreadsheet application. An Excel file is usually called a **workbook**. A workbook contains one or more **worksheets**, which are the tabs you see at the bottom of the file.

This project focuses on moving data between workbooks:

- A **cell** is one box in the worksheet grid.
- A **range** is a rectangular group of cells.
- A **table** is a named range with headers. In VBA, Excel tables are represented as `ListObject` objects.
- A **query** is a saved instruction that fetches or transforms data. Excel often uses **Power Query** for this.
- An **M formula** is the Power Query language used to describe those fetch/transform steps.
- A **connection** is Excel's stored link to an external source, such as another workbook.
- The **Data Model** is Excel's internal storage area for model-backed data and relationships.
- A **macro** is VBA code that automates Excel actions.

The purpose of this sandbox is to make those hidden objects visible: queries, connections, imported tables, and the sometimes surprising metadata Excel keeps after operations are deleted.

---

# Tutorial: first run

This section gives one complete path through the sandbox.

## 1. Prepare a safe environment

Use a disposable workbook or a copy of `SandBox1.xlsm`. The macros create, delete, and rename queries, connections, tables, worksheets, and named ranges.

Recommended prerequisites:

- Microsoft Excel desktop, preferably Excel for Windows.
- Macros enabled for this workbook.
- Access to the OLE DB providers used by the sandbox:
  - `Microsoft.Mashup.OleDb.1`
  - `Microsoft.ACE.OLEDB.12.0`

## 2. Configure the sandbox

Run:

```text
ChangeSettings
```

This walks through the main runtime settings:

- external workbook path
- provider mode
- command type
- target sheet
- target table/range
- default query name
- default connection name

The default values are stored in `ThisWorkbook.cls`. Several macros call `ResetSettings` first so empty settings are restored before the macro continues.

## 3. Create a sample external workbook

Run:

```text
ExportAsDistant
```

This creates a workbook at the configured external path. It builds the configured worksheets and writes the configured column headers into each one. It also names the first-row range of each worksheet as `<sheetName>Range`.

## 4. Create a query

Run:

```text
DoQueries
```

This builds a Power Query/M formula that reads the configured sheet from the external workbook, promotes the first row to headers, and treats configured columns as text.

Depending on what already exists, it asks whether to remove, replace, or add queries.

## 5. Create a connection

Choose one of these:

```text
AddConnectionWithDataModel
```

or:

```text
AddConnectionOnly
```

`AddConnectionWithDataModel` uses `ActiveWorkbook.Connections.Add2` and loads the connection into the Data Model immediately.

`AddConnectionOnly` uses `ActiveWorkbook.Connections.Add`. It can then optionally add the connection to the Data Model through `ActiveWorkbook.Model.AddConnection`.

## 6. Import the data

Choose one of these:

```text
ImportFromConnection
```

or:

```text
ImportFromExternal
```

`ImportFromConnection` imports through an existing workbook connection and uses `xlSrcModel`.

`ImportFromExternal` creates an import directly from the configured provider/source and uses `xlSrcExternal`.

## 7. Inspect the result

Use Excel's UI:

```text
Data tab > Queries & Connections
Data tab > Existing Connections
```

Or use the sandbox macros:

```text
showConnections
showListobjects
```

To reset parts of the workbook state, use:

```text
cleanConnections
cleanListObjects
```

---

# How-to guides

## Change the source workbook

Run `ChangeSettings`, then enter the full path to the external workbook when prompted.

Use this when you want to test a workbook other than the default sample file.

## Switch between provider modes

Run `ChangeSettings` and answer the provider prompt:

- Choose the Mashup/Power Query route to use `Microsoft.Mashup.OleDb.1` and target a workbook query.
- Choose the direct table route to use `Microsoft.ACE.OLEDB.12.0` and target a table/range in the external workbook.

## Change the command type

Run `ChangeSettings` and set one of the supported command types:

| Value | Excel constant | Meaning |
|---:|---|---|
| `2` | `xlCmdSql` | Treat command text as SQL. |
| `3` | `xlCmdTable` | Treat command text as a table name. |
| `6` | `xlCmdTableCollection` | Treat command text as a table collection. |

If a connection fails with an incompatible command type, the helper `errorCommand` may retry with `xlCmdTable`.

## Create a local Excel table from the active sheet

Run:

```text
CreateListObject
```

This creates a `ListObject` from the active sheet's used range or from the named range `<sheetName>TabRef`. If the named range does not exist, the macro creates it from the used range and retries.

## Remove tables from the active sheet

Run:

```text
DeleteListObject
```

This removes tables from the active sheet while trying to preserve the sheet data and clean related named ranges.

## Run automatic management before an import

Run:

```text
setManagement
```

Then select which routines should be active:

- show connections
- clean connections
- show list objects
- clean list objects

`AutoManagement` executes the selected routines in that order. Import macros can call it as a pre-processing step.

## Debug ListObject and connection relationships

Run:

```text
myWeirdestDebug
```

This prints diagnostic output to the Immediate Window. It is intended to compare how `ListObject.QueryTable.WorkbookConnection` and `ListObject.TableObject.WorkbookConnection` relate to workbook connections for different import source types.

---

# Explanation

## What this sandbox is trying to show

Excel can make the same import look simple in the UI while creating several different objects internally. The same visible table may involve:

- a workbook query,
- a workbook connection,
- a model connection,
- a `QueryTable`,
- a `TableObject`,
- a `ListObject`,
- named ranges,
- metadata in the workbook package.

This sandbox isolates those pieces so you can observe which objects appear, which names Excel generates, which objects can be deleted by VBA, and which objects remain as “phantom” metadata.

## Data flow

The general flow is:

```text
external workbook
  -> provider
  -> query or command text
  -> workbook connection
  -> optional Data Model entry
  -> imported table on a worksheet
```

The sandbox lets you test both direct import and connection-based import.

## Provider modes

### Mashup / Power Query provider

The Mashup route uses:

```text
Microsoft.Mashup.OleDb.1
```

The target is a workbook query. The source workbook path and worksheet name are embedded in an M formula created by `DoQueries`.

Use this route when you want to study Power Query behaviour from VBA.

### ACE OLE DB provider

The ACE route uses:

```text
Microsoft.ACE.OLEDB.12.0
```

The target is a table/range in the external workbook. The command text can be a table name or SQL, depending on the command type.

Use this route when you want to test direct workbook-table access without first creating a Power Query query.

## Why the Data Model matters

Some import paths need the connection to be represented in the Data Model before `xlSrcModel` can load data from it.

The sandbox highlights a key practical difference:

- `Connections.Add` creates a connection.
- `Connections.Add2` can create a connection and load it into the Data Model in the same operation.
- `Model.AddConnection` can add an existing connection to the Data Model, but it may recreate or duplicate connection-related objects.

## Why objects sometimes remain after deletion

Excel stores connection and query metadata in more than one place. As a result, deleting visible objects does not always remove all related metadata.

The existing notes call out behaviours such as:

- `ThisWorkbookDataModel` may remain even after connections are deleted.
- Mashup-created connections may not appear in the same UI panels as ACE-created connections.
- Some objects visible in Excel may not be deletable in the same way through VBA.
- Deleting a query manually and deleting a query through VBA can have different effects on related connections and model tables.

Treat these behaviours as observations to reproduce, not as guaranteed Excel API contracts.

## Naming and indentation

Excel automatically changes names to avoid conflicts. The sandbox also adds its own retry logic for names that cause errors.

Examples to watch:

- duplicate queries,
- duplicate connections,
- imported table names,
- model connection names,
- names containing spaces or forbidden characters,
- generated names such as `Connection`, `Table_ExternalData_1`, `ModelConnection_ExternalData_1`, and `ThisWorkbookDataModel`.

## Global settings

The workbook stores runtime settings in private fields in `ThisWorkbook.cls`, exposed through property getters and setters.

Important consequence: after an unhandled VBA runtime error, global state can be lost. Most operational macros call `ResetSettings` at the beginning to restore missing defaults.

---

# Reference

## Project layout

```text
Advanced SandBox for connections/
├── README.md
├── SandBox1.xlsm
└── modules/
    ├── DebugAndTest.bas
    ├── Functions.bas
    ├── GenerConnection.bas
    ├── GenerImportation.bas
    ├── LetManagement.bas
    ├── ManageByCleaning.bas
    ├── ManageByShowing.bas
    ├── ManageQueries.bas
    ├── ModifySettings.bas
    ├── SwitchToTable.bas
    ├── ThisWorkbook.cls
    └── WorkbookModel.bas
```

## Main macros

| Macro | Module | Purpose |
|---|---|---|
| `ExportAsDistant` | `WorkbookModel.bas` | Builds the sample external workbook from configured sheet/table settings. |
| `CreateListObject` | `SwitchToTable.bas` | Converts the active sheet's used range or named range into an Excel table. |
| `DeleteListObject` | `SwitchToTable.bas` | Deletes tables from the active sheet and cleans related names. |
| `DoQueries` | `ManageQueries.bas` | Creates, replaces, or removes workbook queries using a generated M formula. |
| `AddConnectionOnly` | `GenerConnection.bas` | Creates a workbook connection with `Connections.Add`; optionally adds it to the Data Model. |
| `AddConnectionWithDataModel` | `GenerConnection.bas` | Creates a workbook connection with `Connections.Add2` and `CreateModelConnection:=True`. |
| `SafeAceImportADO` | `GenerConnection.bas` | Uses ADO and ACE OLE DB to run SQL and copy a recordset to the active sheet without creating an Excel connection. |
| `ImportFromConnection` | `GenerImportation.bas` | Imports through an existing workbook connection using `xlSrcModel`. |
| `ImportFromExternal` | `GenerImportation.bas` | Imports directly from the configured provider/source using `xlSrcExternal`. |
| `AutoManagement` | `LetManagement.bas` | Runs selected show/clean routines before another operation. |
| `setManagement` | `LetManagement.bas` | Lets the user choose which management routines are active. |
| `getManagement` | `LetManagement.bas` | Interactively runs management routines. |
| `showConnections` | `ManageByShowing.bas` | Displays workbook connection names. |
| `showListobjects` | `ManageByShowing.bas` | Displays worksheet table names. |
| `cleanConnections` | `ManageByCleaning.bas` | Deletes workbook connections except the persistent Data Model placeholder when present. |
| `cleanListObjects` | `ManageByCleaning.bas` | Deletes tables from all worksheets. |
| `ChangeSettings` | `ModifySettings.bas` | Interactively updates runtime settings. |
| `changeTable` | `ModifySettings.bas` | Updates configured table column names for a selected sheet. |
| `myWeirdestDebug` | `DebugAndTest.bas` | Prints connection/table relationship diagnostics. |

## Helper functions

| Function | Module | Purpose |
|---|---|---|
| `errorCommand` | `Functions.bas` | Handles connection/import command errors and retries with a safer command type when possible. |
| `errorName` | `Functions.bas` | Handles duplicate or invalid names by adding an incrementing suffix. |
| `getPosArray` | `Functions.bas` | Finds the position of a string in an array, case-insensitively. |
| `getIndexLimits` | `Functions.bas` | Reads numeric suffixes from names and returns count, lowest index, and highest index. |
| `getOnlyIndex` | `Functions.bas` | Extracts the first numeric part from a name. |
| `QueryExists` | `Functions.bas` | Checks whether a query exists in a workbook query collection. |

## Runtime settings

The main settings are held in `ThisWorkbook.cls`:

| Setting | Meaning |
|---|---|
| external workbook path | Workbook to export to or import from. |
| external worksheets | Sheet names used when generating and reading source data. |
| external tables/columns | Column headers for each configured sheet. |
| target sheet | Sheet to read from in the external workbook. |
| target table/range | Table or named range to read through ACE OLE DB. |
| query name | Default Power Query query name. |
| connection name | Default workbook connection name. |
| provider flag | Chooses Mashup/Power Query mode or ACE direct-table mode. |
| command type | Excel command type, such as SQL or table. |
| management flags | Turn show/clean routines on or off for `AutoManagement`. |

## Known behaviours and caveats

- This is a sandbox. It is intentionally interactive and exploratory.
- Some macros use message boxes and input boxes rather than a formal UI.
- Macros may delete queries, connections, tables, and named ranges.
- The Data Model placeholder `ThisWorkbookDataModel` may remain visible internally even when it appears empty.
- Manual Excel actions and VBA actions may update different internal metadata.
- Provider availability depends on the local Excel/Office installation.
- Default paths and names should be changed before running the macros on a new machine.

## Troubleshooting

### A provider is not available

Check whether the required provider is installed. If ACE OLE DB is missing, install or repair the Microsoft Access Database Engine / Office components that provide it.

### A connection cannot be added

Try using `xlCmdTable` (`3`) as the command type. Some provider/command combinations fail with other command types.

### A table or connection name fails

The helper `errorName` tries to append suffixes such as `_1`, `_2`, and so on. If this still fails, check for invalid characters, duplicate names, or names that begin with numbers.

### Global settings seem empty

Run:

```text
ResetSettings
```

Unhandled VBA errors can clear module-level/global state.

### The Data Model connection remains after cleaning

This is an expected observation in the sandbox. `ThisWorkbookDataModel` can persist as an internal placeholder even after visible connections have been removed.

---

# Documentation structure

This README follows the Diátaxis documentation pattern:

- **Tutorial**: “first run” gives a complete path through the sandbox.
- **How-to guides**: task-focused recipes for changing settings, importing, cleaning, and debugging.
- **Explanation**: concepts behind providers, Data Model behaviour, metadata, and naming.
- **Reference**: project layout, macro list, helper function list, settings, and caveats.

