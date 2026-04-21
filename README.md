# OLAP PivotTable Extensions

> **Active fork** of [OlapPivotTableExtensions/OlapPivotTableExtensions](https://github.com/OlapPivotTableExtensions/OlapPivotTableExtensions) (last upstream release: September 2022).
> This fork fixes broken functionality and modernizes the project.

A C# Excel VSTO add-in (.NET 4.8) that extends PivotTables connected to OLAP/tabular sources: SQL Server Analysis Services, Azure Analysis Services, Power BI, and Power Pivot.

**Current version: 0.9.8**

## What this fork fixes / adds

- **MDX Formatter**: replaced the dead `com.msftlabs.formatmdx` web service with a local `MdxFormatter.cs` implementation — MDX formatting now works offline and without dependency on an external service that no longer exists
- **NLog logging**: structured file logging to `%LOCALAPPDATA%\OlapPivotTableExtensions\logs\` for easier troubleshooting
- **.NET 4.8**: upgraded from .NET 4.5.2
- **Updated packages**: MSAL 4.46 → 4.66, Microsoft.IdentityModel.Abstractions 6.22 → 7.7
- **Bug fixes**: cached `PivotField.Hidden` to avoid 0x800A01A8 COM errors, fixed registry key leaks (`using` statements), fixed `Marshal.ReleaseComObject` misuse on VSTO-managed RCW, replaced silent `catch {}` blocks with logged warnings

## Features

- **Search**: search cube metadata (dimensions, hierarchies, measures, members) directly from Excel
- **Filter & Search**: filter PivotTable fields
- **Calculated Members**: create, save, and load MDX calculated members per workbook
- **MDX Formatting**: format MDX expressions with syntax highlighting
- **KPI support**: display KPI indicators in PivotTables

## Requirements

- Excel 2016 or later
- .NET Framework 4.8

## Build

```
msbuild OlapPivotTableExtensions2016.sln /p:Configuration=Release
```

Available configurations: `Debug`, `Release`, `Release32`, `Release64`.

## Install

Download the latest installer from [Releases](../../releases). The installer registers the add-in in Excel's COM add-in list.

## License

[Microsoft Public License (Ms-PL)](license.md)
