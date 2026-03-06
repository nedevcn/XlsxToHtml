# Nedev.FileConverters.XlsxToHtml

A high-performance, zero-dependency library and CLI for converting Excel (.xlsx) files to HTML targeting **.NET 8.0** (and the library also supports **.NETStandard 2.1**).

## Features

- No third-party dependencies – only .NET Base Class Library.
- Streaming XML parsing for low memory usage.
- Inline CSS styles reflecting cell formatting (fonts, colors, number/date formats; bold, italic, underline, background colors).
- Advanced number formatting: percent, fractions, conditional sections, and custom format strings are interpreted when possible (including basic color codes, positive/negative sections, and hex color specifications).

### Custom colors
A static helper is exposed for named color lookup. Call `ColorHelper.AddOrUpdate("name", "#RRGGBB")` before conversion to register additional mappings (e.g. `"orchid"`).

### Formulas
If a cell contains a formula (`<f>` element) the HTML `<td>` will receive a `title` attribute containing the formula prefixed with `=`; the displayed value normally remains the cached result.  
A runtime option allows basic evaluation of formulas (arithmetic, cell refs, ranges, `SUM`, `AVERAGE`):

```csharp
var writer = new HtmlWriter { EvaluateFormulas = true };
```


- Console utility for batch conversions.

## Usage

Convert a workbook to HTML using the CLI:

```bash
dotnet run --project src/Nedev.FileConverters.XlsxToHtml.Cli -- input.xlsx output.html
```

Omitting the output path will dump the HTML to standard output, making it easy to pipe.

Conversion is streaming and efficient; the entire document is never loaded into a DOM.

## Building & Testing

The library multi-targets .NET 8.0 and .NET Standard 2.1; the CLI and tests target .NET 8.0.

This project depends on the `Nedev.FileConverters.Core` NuGet package. The library
implements `IFileConverter` and is decorated with `[FileConverter("xlsx","html")]`,
enabling automatic discovery by the core infrastructure. The CLI also consumes the
static `Converter.Convert` entry point from the core package.

Use the core API like so:

```csharp
using Nedev.FileConverters;
using var outStream = Converter.Convert(inStream, "xlsx", "html");
```

```bash
# build all projects (run from repo root)
dotnet build src/Nedev.FileConverters.XlsxToHtml.slnx

# run unit tests for the path-based solution
dotnet test src/Nedev.FileConverters.XlsxToHtml.slnx
```

## Limitations

- Formulas are not evaluated by default; set `HtmlWriter.EvaluateFormulas` to `true` for simple expressions.
- Images/charts and complex features (merged cells, comments) are not supported yet.

---

This repository is structured under `src/` following .NET best practices.
