# Security Policy

## Supported versions

OpenXmlMailMerge is pre-release (`0.x`). Only the latest commit on the default branch gets fixes. The public API will
change before `1.0`.

| Version               | Supported |
|-----------------------|-----------|
| `main` / latest `0.x` | Yes       |
| Anything older        | No        |

## Reporting a vulnerability

**Don't open a public issue for security problems.**

Report privately with
GitHub's [private vulnerability reporting](https://docs.github.com/en/code-security/security-advisories/guidance-on-reporting-and-writing-information-about-vulnerabilities/privately-reporting-a-security-vulnerability):
open the repository's **Security** tab and choose **Report a vulnerability**.

Please include:

- the version or commit you tested,
- a minimal template (`.docx`) and the merge data that trigger the problem,
- what happened and what you expected.

This project has a single maintainer, so responses are best effort. You'll get an acknowledgement as soon as possible,
and a fix or mitigation will be coordinated with you before anything is disclosed.

## Threat model

The library reads a `.docx` **template** and **merge data**, and produces a `.docx`. Treat both inputs as untrusted
whenever they come from end users, for example when a web application lets users upload templates. The output is
untrusted for anyone who opens it.

The library doesn't run macros, follow external links, or fetch remote resources. Merged values are inserted as plain
text: markup in a value is written out literally, never interpreted. Parameters read from the template (a table
`STYLE`, the document's language) only select styles and cultures that already exist; an unknown one is ignored. Logs
carry field names and sizes, never merged values.

## Known weaknesses

These are known issues in the current code. Keep them in mind when you deploy the library until they're fixed. Each one
has an ID so issues and commits can refer to it.

| ID     | Weakness                                                                                                                                                                                                           | Impact                                                                                                                                               | Mitigation until fixed                                                                                                                                                     |
|--------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------------------------------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| **S4** | **Few resource limits.** Decompression bombs are stopped by default (`MergeOptions.MaxPartCharacters`), but template size, image size and table row count aren't capped. The whole package is held in memory, plus one copy of the result for `Merge`. | An oversized template, image or table can exhaust memory or CPU.                                                                                     | Enforce size limits before calling the library (for example, a request size limit on uploads). Run merges with a timeout, and isolate them when templates come from users. |
| **S6** | **Image bytes aren't validated.** The content type and size come from the file header, but the image isn't decoded; unrecognized bytes are embedded as `image/jpeg`.                                               | Content that isn't an image is passed through to whoever opens the document. A crafted header can declare any size; it is scaled down to the page width only when the template defines one. | Validate and re-encode images before merging them.                                                                                                                         |
| **S7** | **`MERGEFIELD` instructions of fields without a binding remain in the output.** Bound fields are removed when merged.                                                                                              | The names of unbound fields (part of your data schema) stay visible in the file, and Word can still update them.                                     | Bind every field in the template, or post-process the output to remove the rest.                                                                                           |
| **S8** | **Output isn't validated unless you ask.** The test suite validates every document it produces, but templates it doesn't cover can still produce schema-invalid output.                                            | Consumers may reject the file or parse it unpredictably.                                                                                             | Set `MergeOptions.ValidateOutput`: the merge then throws `MergeException` instead of returning an invalid document.                                                        |

## Fixed

| ID | Weakness                                                                                                                                         | Fixed in                                                            |
|----|--------------------------------------------------------------------------------------------------------------------------------------------------|---------------------------------------------------------------------|
| S1 | Merged documents were written to disk as temp files and left behind by default.                                                                  | In-memory processing; the library no longer writes temp files.      |
| S2 | Temp file names were predictable, and the write wasn't atomic.                                                                                   | Removed with S1.                                                    |
| S3 | Dictionary keys were compiled as regular expressions without a timeout (ReDoS), and values were used as regex replacement patterns (`$0`, `$1`). | Text bindings use plain, ordinal case-insensitive text replacement. |
| S5 | The merge object wasn't thread-safe and kept state between calls, so a shared instance could mix data between requests.                          | `DocumentMerger` keeps no state; each merge has its own private copy of the template. |

XML parsing goes through the Open XML SDK (`DocumentFormat.OpenXml` 3.x) with its default settings. This project has no
tests yet that cover DTD/XXE handling.
