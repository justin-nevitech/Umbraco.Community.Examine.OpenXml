# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Test Commands

```bash
# Build entire solution (all 3 package variants + all 3 test suites + all 3 test sites)
dotnet build src/Umbraco.Community.Examine.OpenXml.slnx

# Build one package variant (v13 | v17 | v18)
dotnet build src/Umbraco.Community.Examine.OpenXml.v18/Umbraco.Community.Examine.OpenXml.v18.csproj

# Run all tests, on every supported Umbraco major
dotnet test src/Umbraco.Community.Examine.OpenXml.slnx

# Run the suite against one major
dotnet test src/Umbraco.Community.Examine.OpenXml.Tests.v18/Umbraco.Community.Examine.OpenXml.Tests.v18.csproj

# Run a single test
dotnet test src/Umbraco.Community.Examine.OpenXml.Tests.v18/Umbraco.Community.Examine.OpenXml.Tests.v18.csproj --filter "FullyQualifiedName~ClassName.MethodName"

# Pack for NuGet (version injected by CI via /p:Version; package major must match the variant)
dotnet pack src/Umbraco.Community.Examine.OpenXml.v18/Umbraco.Community.Examine.OpenXml.v18.csproj -c Release
```

## Architecture

This is an Umbraco CMS package that extracts text from OpenXml documents (.docx, .pptx, .xlsx) in the media library and indexes it into a dedicated Examine/Lucene index called `OpenXmlIndex`.

### Multi-major support

`src/Umbraco.Community.Examine.OpenXml/` is **not a project** — it is the shared source folder,
holding all C# sources plus `Examine.OpenXml.Shared.props`. Three wrapper projects own no sources
of their own, import that props file, and compile the same files against a different Umbraco major:

| Project | TFM | Umbraco.Cms.Web.Common | Produces |
|---|---|---|---|
| `Umbraco.Community.Examine.OpenXml.v13` | net8.0 | `[13.0.0, 14.0.0)` | package `13.x` |
| `Umbraco.Community.Examine.OpenXml.v17` | net10.0 | `[17.0.0, 18.0.0)` | package `17.x` |
| `Umbraco.Community.Examine.OpenXml.v18` | net10.0 | `[18.0.0, 19.0.0)` | package `18.x` |

TFM-based multi-targeting cannot express this because Umbraco 17 and 18 both run on net10.0. All
three variants share one `PackageId` and one assembly name; releases are version-aligned, with the
package major tracking the Umbraco major. Add new sources to the shared folder — every variant
picks them up automatically.

The code is identical across all majors — no `#if` preprocessor directives needed.

**Full detail lives in [docs/BUILDING.md](docs/BUILDING.md). Read it before touching the project
layout, the release workflow, or the supported-version set.**

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

- `src/Umbraco.Community.Examine.OpenXml/` — Shared package sources (**not a project**) + `Examine.OpenXml.Shared.props`
- `src/Umbraco.Community.Examine.OpenXml.v13|.v17|.v18/` — Package projects, one per Umbraco major
- `src/Umbraco.Community.Examine.OpenXml.Tests/` — Shared test sources (**not a project**, xUnit + Moq, 119 tests) + `Examine.OpenXml.Tests.Shared.props`
- `src/Umbraco.Community.Examine.OpenXml.Tests.v13|.v17|.v18/` — Test projects, one per Umbraco major
- `src/Umbraco.Community.Examine.OpenXml.TestSite.v13/` — Umbraco 13 test site (net8.0, port 44380)
- `src/Umbraco.Community.Examine.OpenXml.TestSite.v17/` — Umbraco 17 test site (net10.0, port 44379)
- `src/Umbraco.Community.Examine.OpenXml.TestSite.v18/` — Umbraco 18 test site (net10.0, port 44381)

### Test Sites

Each test site uses the Clean starter kit, uSync for content import, and unattended install (admin@example.com / 1234567890). The v13 site uses uSync folder `v9/`, v17 uses `v17/` and v18 uses `v18/`. Each site references its own matching package project via ProjectReference, and all three can run at once.

Unlike the package projects — which pin the **lowest** release of their major so every patch is a valid host — the test sites track the **latest** release of their major.

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

### Multi-major support
- Pin the lowest release of each Umbraco major (13.0.0, 17.0.0, 18.0.0) as the range floor so all patch versions are compatible
- Put new sources in the shared folders, never in a `.vNN` wrapper project — wrappers own no sources
- Adding or dropping an Umbraco major means: a package project, a test project, a test site, a `release.yml` case, the `.slnx`, and `docs/BUILDING.md`
- Only reference packages actually used — check with `grep` for namespace usage before adding
- Avoid Umbraco API packages (`Umbraco.Cms.Api.Common`, `Umbraco.Cms.Api.Management`) unless the code references their types — they cause version conflicts

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

Releases are version-aligned: the package major tracks the Umbraco major. Push a semantic version
tag whose major is 13, 17 or 18 (e.g. `18.0.0`) to trigger the GitHub Actions workflow, which
selects the matching package project from the tag, runs that major's test suite, then packs and
publishes to NuGet using the `NUGET_API_KEY` secret. A tag with any other major fails the build.
