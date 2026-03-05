# Nedev.XlsxToHtml

A high-performance, zero-dependency .NET 10 library and CLI for converting Excel (.xlsx) files to HTML.

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
dotnet run --project src/Nedev.XlsxToHtml.Cli -- input.xlsx output.html
```

Omitting the output path will dump the HTML to standard output, making it easy to pipe.

Conversion is streaming and efficient; the entire document is never loaded into a DOM.

## Building & Testing

```bash
dotnet build

dotnet test
```

## Limitations

- Formulas are not evaluated by default; set `HtmlWriter.EvaluateFormulas` to `true` for simple expressions.
- Images/charts and complex features (merged cells, comments) are not supported yet.

---

This repository is structured under `src/` following .NET best practices.
