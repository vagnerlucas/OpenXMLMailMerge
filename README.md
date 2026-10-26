# OpenXmlMailMerge

[![CI](https://github.com/vagnerlucas/OpenXMLMailMerge/actions/workflows/ci.yml/badge.svg?branch=main)](https://github.com/vagnerlucas/OpenXMLMailMerge/actions/workflows/ci.yml?query=branch%3Amain)

Merge data into Word (`.docx`) templates. You design the template in Microsoft Word using `MERGEFIELD` fields, and the
library replaces those fields with text, images or tables. It's built on
the [Open XML SDK](https://github.com/dotnet/Open-XML-SDK), so Word doesn't need to be installed.

> **Status: pre-release (`0.x`).** The public API may still change before `1.0`, and the current version has known
> limitations (see below). It isn't published on NuGet yet.

## What it does

| Binding | You supply                                                  | The field becomes                                                          |
|---------|-------------------------------------------------------------|----------------------------------------------------------------------------|
| `Text`  | Any value (text, numbers, dates)                            | The value as text, formatted in the template's language                    |
| `Image` | The bytes of a PNG, JPEG, GIF, BMP, TIFF, EMF or WMF        | An inline picture at the image's own size                                  |
| `Table` | A `TableValue` (or a `DataTable` through `TableValue.From`) | A table with a header row and one row per data row, styled by the template |

Fields are processed in the document body, headers and footers.

## Requirements

- [.NET 10 SDK](https://dotnet.microsoft.com/download) to build (pinned in `global.json`).
- The library targets `net10.0`, so applications that use it need .NET 10.

## Build and test

```sh
dotnet build
dotnet test
```

Every test that merges a document also validates it against the Open XML schema. The merged documents are kept in
`tests/OpenXmlMailMerge.Tests/bin/Debug/net10.0/Output/`, one per test, so you can open them in Word. CI runs the tests
on Linux, Windows and macOS, and attaches those documents to each run.

The public API is recorded in `src/OpenXmlMailMerge/PublicAPI.Shipped.txt` and `PublicAPI.Unshipped.txt`, nullable
annotations included. A public change that isn't recorded there fails the build (CI builds with warnings as errors), so an
API change is always deliberate. Record a new member in `PublicAPI.Unshipped.txt`; the IDE's quick fix for `RS0016` writes
the line.

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
`«CustomerName»` but not `«CustomerNameFull»`. A bound field is replaced at its position (text and images keep the
field's formatting), and its field code is removed. Fields without a binding are left as they are. Fields saved by Word
as simple fields (`w:fldSimple`) and as complex fields (`w:fldChar`, spread over several runs) are both merged.

### Field parameters

Images and tables accept parameters after the field name, each starting with `@`. They aren't part of the name, so the
key `Logo` matches the field `Logo@WIDTH=120`. Parameter names ignore case.

| Parameter         | Applies to | Example                    |
|-------------------|------------|----------------------------|
| `WIDTH`, `HEIGHT` | Images     | `Logo@WIDTH=120@HEIGHT=80` |
| `STYLE`           | Tables     | `Items@STYLE=GridTable4`   |

To add parameters to a field already in the template, show the field codes in Word (**Alt+F9**, or **Option+F9** on a
Mac), edit the text after `MERGEFIELD`, then press the same keys to hide the codes again.

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
change the result. `MergeOptions.Culture` overrides it.

## Usage

```csharp
using OpenXmlMailMerge;

var data = new MergeData()
    .Text("CustomerName", "Jane Doe")
    .Text("Total", 1234.5m)                               // formatted in the template's language
    .Image("Logo", File.ReadAllBytes("logo.png"))
    .Table("Items", TableValue.From(itemsDataTable));     // or new TableValue(columns, rows)

var merger = new DocumentMerger();                        // reusable and thread-safe
MergeResult result = merger.Merge(File.ReadAllBytes("template.docx"), data);
File.WriteAllBytes("result.docx", result.Document);
```

With streams, for example in ASP.NET Core, use the async overload. It reads the template to its end, writes the merged
document to the output, and closes neither stream:

```csharp
MergeReport report = await merger.MergeAsync(templateStream, Response.Body, data, cancellationToken);
```

The merge runs entirely in memory, on a private copy of the template; the template is never modified. Only
`MergeAsync` is async, because only it does I/O.

A `DocumentMerger` keeps nothing between merges, so one instance can serve a whole application, for example as a
singleton in dependency injection. A `MergeData` can be merged into any number of templates.

The output is deterministic: the same template and data always produce the same bytes, so results can be cached, hashed
or compared in tests.

### The report

Every merge returns a report of what didn't match, to catch a typo in the template or in your code:

- `UnboundFields`: fields that got no data, for example `«Adress»` when the data says `Address`. Bound fields that can't
  be merged (nested in another field, or spread over paragraphs) are listed too, and so are fields that
  `UnboundFieldBehavior.Remove` took out.
- `UnusedBindings`: names in your data that no field used.

These are information, not errors, unless you set `UnboundFields = UnboundFieldBehavior.Throw`. Data without any binding
is valid: the template comes back with its fields untouched.

### Options

```csharp
var merger = new DocumentMerger(new MergeOptions
{
    UnboundFields = UnboundFieldBehavior.Remove,
    Culture = CultureInfo.GetCultureInfo("en-US"),
});
```

| Option              | Default             | Effect                                                                                                                                                                      |
|---------------------|---------------------|-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| `UnboundFields`     | `Keep`              | What happens to a field without data. `Keep` leaves its placeholder (`«Name»`). `Remove` deletes the field and its placeholder. `Throw` fails the merge and writes nothing. |
| `Culture`           | `null` (template's) | The culture numbers and dates are formatted with.                                                                                                                           |
| `MaxPartCharacters` | `50000000`          | The largest any XML part of the template may grow to once decompressed. Stops decompression bombs: small templates whose XML inflates to gigabytes. `0` disables the check. |
| `ValidateOutput`    | `false`             | Validate the merged document against the Open XML schema, and throw `MergeException` instead of returning an invalid one. Slower, because it reads the whole document.      |

`Remove` can't safely take out a field nested in another field or spread over paragraphs; those stay, and the report
lists them.

### Inspecting a template

`Inspect` lists a template's merge fields without merging anything, for example to check a template before using it, or
to show its fields in a UI:

```csharp
foreach (var field in merger.Inspect(File.ReadAllBytes("template.docx")).Fields)
    Console.WriteLine($"{field.Name} in {field.PartUri}, mergeable: {field.IsMergeable}");
```

Each field comes with its parameters (`WIDTH`, `STYLE`, ...), whether Word stored it as a simple or a complex field, and
whether a merge can fill it: a field nested in another field or spread over paragraphs can't be.

### Logging

`DocumentMerger` logs through [Microsoft.Extensions.Logging](https://learn.microsoft.com/dotnet/core/extensions/logging),
so logs go wherever your application sends them: the console, OpenTelemetry (and from there New Relic, Datadog, ...),
Application Insights, CloudWatch, Serilog sinks, and so on. With dependency injection the logger is passed in for you;
otherwise pass one:

```csharp
using var loggerFactory = LoggerFactory.Create(builder => builder.AddConsole().SetMinimumLevel(LogLevel.Debug));
var merger = new DocumentMerger(logger: loggerFactory.CreateLogger<DocumentMerger>());
```

Without a logger, nothing is logged. Merges log at `Debug` (start, duration, size, unbound fields and unused bindings),
plus a `Warning` when a binding matches a field that can't be merged. Logs carry field names and sizes, never merged
values.

### Errors

Invalid arguments throw right away: an empty or duplicate name (names ignore case), an empty image, a table row with the
wrong number of values, or an option out of range. Problems with the template throw `MergeException`:

- the template isn't a valid `.docx`;
- a part is larger than `MaxPartCharacters` (`PartUri` names it);
- fields got no data under `UnboundFieldBehavior.Throw` (`FieldName` and `PartUri` name the first one);
- the result isn't valid under `ValidateOutput` (`PartUri` names the part).

## Known limitations

- **Nested and multi-paragraph fields are left as they are.** A merge field nested in another field
  (`{ IF { MERGEFIELD Name } ... }`) or spread over more than one paragraph isn't merged.
- **Footnotes and endnotes aren't merged.** Only the body, headers and footers are.
- **Some images get a default size of 42×42 px** unless the field name sets one: a WMF without the placeable header
  (it carries no physical size), and any image whose header can't be read. Bytes in a format the library doesn't
  recognize are embedded as they are and labelled `image/jpeg`, which Word can't display.
- **Image resolution (DPI) is ignored.** Raster images are sized at 96 DPI, so a 300 DPI scan shows larger than it would
  in Word.
- **Tables are generated, not designed in the template.** You choose their text formatting and table style, but you
  can't lay out a row in Word and have it repeated yet. A table style must be defined in the document: Word only stores
  the styles a document uses, so apply the style once in the template (to any table) before naming it in `STYLE`.
- **The whole package is held in memory** during a merge, and `Merge` adds one copy of the result (`MergeAsync` writes
  straight from the package instead). See [docs/SECURITY.md](docs/SECURITY.md) for limits to apply to untrusted
  templates.

## Roadmap

- More limits for untrusted input: template, image and table sizes ([S4](docs/SECURITY.md))
- Repeating regions (`TableStart:Name` / `TableEnd:Name`) that reuse a styled template row
- Merge fields nested in other fields (`IF`, `INCLUDETEXT`, ...), footnotes and endnotes
- Image resolution (DPI)
- Golden-file tests

## How this refactor was made

OpenXmlMailMerge started in 2016 as a .NET Framework 4.5 library. In 2026 I brought it back and refactored it with
[Claude](https://claude.ai), Anthropic's AI model, working in Claude Code. I set the direction and made edits of my own
along the way. Claude reviewed the old code, proposed the plan, and wrote most of the changes and tests.

In short, the project moved to modern .NET with an SDK-style layout and gained:

- in-memory processing with async I/O;
- exact field-name matching, for simple and complex fields;
- images and tables placed where their field is, with images at their own size in every supported format;
- tables and values that follow the template's formatting, styles and language;
- a new, stateless public API (`DocumentMerger`) with options, a merge report and template inspection, its surface
  tracked by an analyzer;
- a test suite that validates every document it produces, with CI on Linux, Windows and macOS.

> *A note from Claude:* every bug fixed here was reproduced first, and the new tests were run against the old code to
> show what they catch. What isn't fixed yet is listed under Known limitations, not hidden.
> — Claude (Opus 5.5), Anthropic

## Security

Please report vulnerabilities privately. See [docs/SECURITY.md](docs/SECURITY.md).

## License

[MIT](LICENSE) © 2016–2026 Vagner Lucas

The sample images used by the tests come from other projects under their own licenses (public domain, CC0, libtiff and
Apache-2.0). See [tests/OpenXmlMailMerge.Tests/Fixtures/Images/README.md](tests/OpenXmlMailMerge.Tests/Fixtures/Images/README.md).
