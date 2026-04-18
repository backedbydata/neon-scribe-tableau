# NeonScribe for Tableau

A Windows desktop application that automatically documents Tableau dashboards from TWB and TWBX files, generating clean HTML documentation focused on the end-user perspective — what the dashboard shows, how filters work, and what actions do — rather than internal implementation details.

## Features

- Parses TWB and TWBX files (both packaged and unpackaged workbooks)
- Resolves Tableau's internal cryptic field names to human-readable captions
- Documents worksheets, dashboards, filters, parameters, calculated fields, and actions
- Generates styled HTML output or JSON for programmatic use
- WPF desktop UI and CLI both included

## Project Structure

```
NeonScribe.Tableau/
├── src/
│   ├── NeonScribe.Tableau.Core/           # Parser and object model
│   ├── NeonScribe.Tableau.Documentation/  # HTML/JSON output generation
│   ├── NeonScribe.Tableau.CLI/            # Command-line interface
│   └── NeonScribe.Tableau.WPF/            # Windows desktop UI
├── samples/                               # Real-world Tableau workbooks for testing
│   ├── Combine/
│   ├── Hospital/
│   ├── Map/
│   ├── MultiFilterCustomButtons/
│   ├── PeriodicTable/
│   ├── Super/
│   └── TableHeaders/
├── scripts/                               # Developer utility scripts
│   ├── analyze-twb.ps1                    # Analyze TWB complexity metrics
│   ├── download-samples.ps1               # Download sample workbooks
│   └── test-comparison.ps1                # Compare parser output
└── assets/                                # Application icons and images
```

## Prerequisites

- Windows 10 or later
- .NET 10.0 SDK or later
- PowerShell 5.1 or later (for utility scripts)

## Quick Start

### Build

```bash
git clone https://github.com/backedbydata/neon-scribe-tableau.git
cd neon-scribe-tableau
dotnet build NeonScribe.Tableau.slnx
```

### WPF Desktop App

```powershell
.\build-and-run-wpf.ps1
```

### CLI Usage

```bash
# Generate HTML documentation
dotnet run --project src/NeonScribe.Tableau.CLI -- your-workbook.twb -o documentation.html -f html

# Generate JSON output
dotnet run --project src/NeonScribe.Tableau.CLI -- your-workbook.twb -o output.json -f json

# Example with included samples
dotnet run --project src/NeonScribe.Tableau.CLI -- samples/Hospital/"#RWFD Hospital ER Dashboard.twb" -o hospital-docs.html -f html
```

## What Gets Documented

| Section | Details |
|---|---|
| **Dashboards** | Layout, zones, embedded worksheets |
| **Worksheets** | Visual type, fields used, mark configuration |
| **Filters & Parameters** | Label, control type, default values, multi-select |
| **Calculated Fields** | Formula, dependencies, usage across sheets |
| **Actions** | Type (filter/highlight/URL/parameter), trigger, source/target |
| **Tooltips** | Content and fields displayed |
| **Data Sources** | Connections, field lineage |

## Scripts

### Analyze TWB Complexity

```powershell
# Single file
.\scripts\analyze-twb.ps1 -Path "samples\Hospital\#RWFD Hospital ER Dashboard.twb"

# All samples, export to CSV
.\scripts\analyze-twb.ps1 -Path "samples" -Recursive -OutputCsv "analysis.csv"
```

Reports worksheets, dashboards, field counts, calculation types, filter counts, actions, and a 0–100 complexity score.

### Download Additional Samples

```powershell
.\scripts\download-samples.ps1
```

Downloads real-world Tableau workbooks from public GitHub repositories into `samples/`.

## Technical Notes

**TWB files** are XML containing workbook metadata — data sources, worksheets, dashboards, calculations, filters, parameters, and actions — but no actual data.

**TWBX files** are ZIP archives containing a TWB plus data extracts (`.hyper`/`.tde`), images, and other resources. The parser extracts and processes these automatically.

**Name resolution:** Tableau internally uses cryptic IDs like `[federated.xyz].[calculation_123]`. NeonScribe resolves these to human-readable captions using the `caption` attribute and `<aliases>` sections in the XML.

## License

MIT — see [LICENSE](LICENSE) for details.

## Contributing

Contributions, bug reports, and feature requests are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.
