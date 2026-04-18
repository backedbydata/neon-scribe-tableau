# Contributing to NeonScribe for Tableau

Thanks for your interest in contributing. Here's how to get started.

## Prerequisites

- Windows 10 or later
- .NET 10.0 SDK
- Visual Studio 2022 or VS Code with the C# extension

## Getting Started

```bash
git clone https://github.com/backedbydata/neon-scribe-tableau.git
cd neon-scribe-tableau
dotnet build NeonScribe.Tableau.slnx
```

## Project Structure

| Project | Purpose |
|---|---|
| `NeonScribe.Tableau.Core` | XML parsing and object model |
| `NeonScribe.Tableau.Documentation` | HTML/JSON output generation |
| `NeonScribe.Tableau.CLI` | Command-line interface |
| `NeonScribe.Tableau.WPF` | Windows desktop UI |

## Making Changes

1. Fork the repository and create a branch from `main`
2. Make your changes in the appropriate project
3. Test against the sample workbooks in `samples/`
4. Verify the CLI produces valid output: `dotnet run --project src/NeonScribe.Tableau.CLI -- samples/Hospital/"#RWFD Hospital ER Dashboard.twb" -o test.html -f html`
5. Open a pull request with a clear description of what changed and why

## Reporting Bugs

Open a GitHub issue with:
- The TWB/TWBX file if you're able to share it (or a minimal reproduction)
- The exact command or steps to reproduce
- The actual output vs. what you expected

## Areas for Contribution

- Support for additional Tableau XML features (stories, level-of-detail edge cases)
- Improved HTML output styling
- PDF export
- Additional output formats (Markdown, DOCX)
- Better handling of very large or complex workbooks

## License

By contributing, you agree that your contributions will be licensed under the MIT License.
