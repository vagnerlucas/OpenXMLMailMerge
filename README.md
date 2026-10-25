# OpenXmlMailMerge

[![CI](https://github.com/vagnerlucas/OpenXMLMailMerge/actions/workflows/ci.yml/badge.svg)](https://github.com/vagnerlucas/OpenXMLMailMerge/actions/workflows/ci.yml)

Merge data into Word (`.docx`) templates. You design the template in Microsoft Word using `MERGEFIELD` fields, and the
library replaces those fields with text, images or tables. It's built on
the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK), so Word doesn't need to be installed.

> **Status: pre-release (`0.x`).** The engine is being rewritten. The public API will change before `1.0`, and the
> current version has known limitations (see below). It isn't published on NuGet yet.

## What it does

| Binding | You supply              | The field becomes                                  |
|---------|-------------------------|----------------------------------------------------|
| `Text`  | `string`                | The text value                                     |
| `Image` | `byte[]`                | An inline picture                                  |
| `Table` | `System.Data.DataTable` | A table with a header row and one row per data row |

Fields are processed in the document body, headers and footers.

## Requirements

- [.NET 10 SDK](https://dotnet.microsoft.com/download) (pinned in `global.json`)

## Build and test

```sh
dotnet build
dotnet test
```

To produce a local NuGet package:

```sh
dotnet pack src/OpenXmlMailMerge -c Release -o artifacts
```

## Writing a template

1. In Word, choose **Insert → Quick Parts → Field… → MergeField**.
2. Enter the field name, for example `CustomerName`. Word shows it as `«CustomerName»`.
3. Save the document as `.docx`.

`tests/OpenXmlMailMerge.Tests/Fixtures/BIND.docx` is a working example with fields in the body, header and footer.

A binding key matches a field when it equals the field's name exactly, ignoring case: `CustomerName` matches
`«CustomerName»` but not `«CustomerNameFull»`. For images, parameters after `@` aren't part of the name, so `Logo` matches
`Logo@WIDTH=120`. A bound field is replaced at its position (text and images keep the field's formatting), and its field
code is removed. Fields without a binding are left as they are. Fields saved by Word as simple fields (`w:fldSimple`) and
as complex fields (`w:fldChar`, spread over several runs) are both merged.

### Images

An image keeps its own size, scaled down to the width between the page margins. This applies in the body, headers and
footers. For PNG, JPEG, GIF, BMP and TIFF that's the pixel size at 96 DPI; for EMF and WMF it's the physical size in the
file's header. To set the size, add it to the field name in pixels (`Logo@WIDTH=120@HEIGHT=80`), or give just one
dimension to keep the aspect ratio. PNG, JPEG, GIF, BMP, TIFF, EMF and WMF are stored with their own content type.

### Tables

A table field becomes a table with a header row (the column captions, repeated on each page) and one row per data row.
It replaces the field's paragraph when the field is alone in it, and goes right after the paragraph otherwise. The
template decides how it looks:

- **Text:** cells take the field's formatting (font, size, bold, ...). Anything the field doesn't set comes from the
  document's defaults.
- **Table style:** add `STYLE` to the field name with a table style id or name, for example `Items@STYLE=GridTable4`.
  Quote a name with spaces: `MERGEFIELD "Items@STYLE=Grid Table 4"`. Without `STYLE`, or when the document doesn't
  define that style, the table gets single borders.

### Numbers and dates

Values merged into text fields and table cells are formatted in the template's language: the language set on the
field's text, else the document's default language, else the invariant culture. The machine running the merge doesn't
change the result.

## Usage

```csharp
using OpenXmlMailMerge;

var merge = new OxMailMerge(DocumentType.Doc);
merge.Document.AddToDictionary(MailMergeDataTypeEnum.Text, "CustomerName", "Jane Doe");
merge.Document.AddToDictionary(MailMergeDataTypeEnum.Image, "Logo", File.ReadAllBytes("logo.png"));
merge.Document.AddToDictionary(MailMergeDataTypeEnum.Table, "Items", itemsDataTable);

try
{
    await merge.Document.LoadFromFileAsync("template.docx");
    merge.Process();
    await merge.Document.SaveToFileAsync("result.docx");
}
finally
{
    merge.Terminate();                    // releases the in-memory package
}
```

The merge runs entirely in memory: loading copies the template into a private buffer, and the template is never
modified. Load and save the way that fits your code:

| Source / destination                                                  | Load                              | Save                            |
|-----------------------------------------------------------------------|-----------------------------------|---------------------------------|
| `byte[]`                                                              | `LoadFromBytes(bytes)`            | `GetBytes()`                    |
| `Stream` (for example, ASP.NET Core `Request.Body` / `Response.Body`) | `LoadFromStreamAsync(stream, ct)` | `SaveToStreamAsync(stream, ct)` |
| File                                                                  | `LoadFromFileAsync(path, ct)`     | `SaveToFileAsync(path, ct)`     |

Only methods that do I/O are async. `Process()` is synchronous because it's CPU-bound work on an in-memory document. You
can save more than once; each save includes every change made so far.

Use one `OxMailMerge` instance per document. Instances aren't thread-safe.

## Known limitations

- **Nested and multi-paragraph fields are left as they are.** A merge field nested in another field
  (`{ IF { MERGEFIELD Name } ... }`) or spread over more than one paragraph isn't merged.
- **Some images get a default size of 42×42 px** unless the field name sets one: a WMF without the placeable header
  (it carries no physical size), and any image whose header can't be read. Bytes in a format the library doesn't
  recognize are embedded as they are and labelled `image/jpeg`, which Word can't display.
- **Image resolution (DPI) is ignored.** Raster images are sized at 96 DPI, so a 300 DPI scan shows larger than it would
  in Word.
- **Tables are generated, not designed in the template.** You choose their text formatting and table style, but you
  can't lay out a row in Word and have it repeated yet. A table style must be defined in the document: Word only stores
  the styles a document uses, so apply the style once in the template (to any table) before naming it in `STYLE`.
- **The whole package is held in memory**, plus one copy for each save. See [docs/SECURITY.md](docs/SECURITY.md) for
  limits to apply to untrusted templates.

## Roadmap

- A stateless API: `Merge(template, data) → document`
- Merge fields nested in other fields (`IF`, `INCLUDETEXT`, ...)
- Deterministic ids for images and relationships
- Repeating regions (`TableStart:Name` / `TableEnd:Name`) that reuse a styled template row
- Golden-file tests and an option to validate output at runtime

## How this refactor was made

OpenXmlMailMerge started in 2016 as a .NET Framework 4.5 library. In 2026 I brought it back and refactored it with
[Claude](https://claude.ai), Anthropic's AI model, working in Claude Code. I set the direction and made edits of my own
along the way. Claude reviewed the old code, proposed the plan, and wrote most of the changes and tests.

In short, the project moved to modern .NET with an SDK-style layout and gained:

- in-memory processing with async I/O;
- exact field-name matching;
- images and tables placed and sized correctly;
- a smaller public API;
- a test suite that validates every document it produces.

> *A note from Claude:* every bug fixed here was reproduced first, and the new tests were run against the old code to
> show what they catch. What isn't fixed yet is listed under Known limitations, not hidden. Thanks for asking for blunt
> reviews, Vagner, and for pushing back when something looked off.
> — Claude (Opus 5.5), Anthropic

## Security

Please report vulnerabilities privately. See [docs/SECURITY.md](docs/SECURITY.md).

## License

[MIT](LICENSE) © 2016–2026 Vagner Lucas
