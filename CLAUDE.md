# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Test Commands

```bash
# Build entire solution
dotnet build src/Umbraco.Community.Examine.OpenXml.slnx

# Build package project only (all 3 TFMs)
dotnet build src/Umbraco.Community.Examine.OpenXml/Umbraco.Community.Examine.OpenXml.csproj

# Run all tests
dotnet test src/Umbraco.Community.Examine.OpenXml.Tests/Umbraco.Community.Examine.OpenXml.Tests.csproj

# Run a single test
dotnet test src/Umbraco.Community.Examine.OpenXml.Tests/Umbraco.Community.Examine.OpenXml.Tests.csproj --filter "FullyQualifiedName~ClassName.MethodName"

# Pack for NuGet (version injected by CI via /p:Version)
dotnet pack src/Umbraco.Community.Examine.OpenXml/Umbraco.Community.Examine.OpenXml.csproj -c Release
```

## Architecture

This is an Umbraco CMS package that extracts text from OpenXml documents (.docx, .pptx, .xlsx) in the media library and indexes it into a dedicated Examine/Lucene index called `OpenXmlIndex`.

### Multi-targeting

The package targets `net8.0`, `net9.0`, and `net10.0` with conditional Umbraco package references:
- net8.0 → Umbraco.Cms.Web.Common 13.0.0
- net9.0 → Umbraco.Cms.Web.Common 16.0.0
- net10.0 → Umbraco.Cms.Web.Common 17.0.0

The code is identical across all targets — no `#if` preprocessor directives needed.

### Core Flow

1. **Registration**: `ExamineOpenXmlComposer` (IComposer) calls `AddExamineOpenXml()` which registers all services and the Lucene index via DI.

2. **Indexing on media change**: `OpenXmlCacheNotificationHandler` listens for `MediaCacheRefresherNotification`, checks if media is a supported OpenXml type, and calls `OpenXmlIndexPopulator.AddToIndex()` or `RemoveFromIndex()` for create/update/delete/trash events.

3. **Full index rebuild**: `OpenXmlIndexPopulator.PopulateIndexes()` pages through all media, filters by file extension, and builds value sets via `OpenXmlIndexValueSetBuilder`.

4. **Text extraction chain**: `OpenXmlService` → `OpenXmlTextExtractorFactory` (routes by extension) → specific extractor (`WordProcessingDocumentTextExtractor`, `PresentationDocumentTextExtractor`, or `SpreadsheetDocumentTextExtractor`).

5. **Index validation**: `OpenXmlValueSetValidator` filters out items in the recycle bin and validates parent ID paths during indexing.

### Text Extractors

- **Word**: Uses `Paragraph.InnerText` on body, headers, footers, footnotes, endnotes. This correctly handles Word's split-run behavior where a single word spans multiple `<w:r>` elements.
- **PowerPoint**: Iterates `Drawing.Paragraph` descendants on each slide and its notes slide.
- **Spreadsheet**: Reads cells via `OpenXmlReader`, resolves shared strings from `SharedStringTablePart` using a pre-materialized list for O(1) lookup.

All extractors wrap OpenXml documents in `using` statements to prevent resource leaks.

### Key Constants (OpenXmlIndexConstants)

- Index name: `"OpenXmlIndex"`
- Content field: `"fileTextContent"`
- Category: `"openxml"`
- Supported extensions: `"docx"`, `"pptx"`, `"xlsx"`
- Max file size: 100 MB — files exceeding this are skipped before parsing
- Max extracted content length: 10 MB — extraction stops at this limit
- Max characters per part: 10,000,000 — limits per-part decompression via `OpenSettings.MaxCharactersInPart`
- Max shared string count: 1,000,000 — caps Excel shared string table materialization

## Solution Structure

- `src/Umbraco.Community.Examine.OpenXml/` — Package library (multi-targeted)
- `src/Umbraco.Community.Examine.OpenXml.Tests/` — Unit tests (xUnit + Moq, 119 tests)
- `src/Umbraco.Community.Examine.OpenXml.TestSite/` — Umbraco 17 test site (net10.0, port 44379)
- `src/Umbraco.Community.Examine.OpenXml.TestSite.v13/` — Umbraco 13 test site (net8.0, port 44380)
- `src/Umbraco.Community.Examine.OpenXml.TestSite.v16/` — Umbraco 16 test site (net9.0, port 44381)

### Test Sites

Each test site uses the Clean starter kit, uSync for content import, and unattended install (admin@example.com / 1234567890). The v13 site uses uSync folder `v9/`, while v16 and v17 use `v17/`. All three test sites reference the package via ProjectReference.

## Coding Standards

### Umbraco Package Conventions
- Register services via `IComposer` + `IUmbracoBuilder` extension methods — no manual `Program.cs` changes for consumers
- Reference `Umbraco.Cms.Web.Common` (not the full `Umbraco.Cms` meta-package) to minimize dependency footprint
- Use `AddUnique` for service registrations that consumers might want to override
- Notification handlers must never throw — log and return gracefully to avoid breaking the Umbraco pipeline
- Use `IRuntimeState.Level` checks to skip processing during install/upgrade

### Examine Index Conventions
- Custom indexes inherit from `LuceneIndex` with `IIndexDiagnostics` for the backoffice Examine Management dashboard
- Use `IndexPopulator` for full rebuilds and notification handlers for incremental updates
- `ValueSetValidator` must filter recycle bin items to prevent trashed content appearing in search
- Always handle all media lifecycle events: create, update, delete, trash, restore, branch moves

### .NET / C# Standards
- All `IDisposable` objects (OpenXml documents, streams, readers) must be in `using` statements
- Use `int.TryParse` instead of `int.Parse` for external input (file content, user data)
- Use specific exception types (`NotSupportedException`, `InvalidOperationException`) not generic `Exception`
- Don't include user-supplied values in exception messages (information disclosure risk)
- Add content size limits when processing untrusted files to prevent OOM from malicious documents
- Null-coalesce (`?? string.Empty`) when passing nullable values to methods that don't accept null (e.g. `Contains()`)

### Multi-targeting
- Use lowest minor version per Umbraco major (13.0.0, 16.0.0, 17.0.0) so all patch versions are compatible
- Only reference packages actually used — check with `grep` for namespace usage before adding
- Avoid Umbraco API packages (`Umbraco.Cms.Api.Common`, `Umbraco.Cms.Api.Management`) unless the code references their types — they cause version conflicts in multi-target builds

### Testing
- Unit tests must not require an Umbraco instance or Lucene index
- Use real OpenXml documents (from `samples/`) for extractor tests, not mocks
- Use `DocumentFormat.OpenXml` to create in-memory documents for edge case tests
- Use Moq for Umbraco service dependencies (`IMediaService`, `IExamineManager`, etc.)

## Custom Commands

Project-specific slash commands in `.claude/commands/`:

| Command | Use For |
|---------|---------|
| `/project:review` | Code quality review against project standards (resource leaks, null safety, index consistency, text extraction completeness) |
| `/project:pre-publish` | Full pre-publish checklist — build, test, pack, inspect nupkg, verify docs and CI/CD |
| `/project:security-scan` | Security review — dependency CVEs, input validation, resource safety, XSS/injection checks |

## Release Process

Push a semantic version tag (e.g. `1.0.0`) to trigger the GitHub Actions workflow which packs and publishes to NuGet using the `NUGET_API_KEY` secret.
