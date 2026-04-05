# Research

Technical decisions and rationale behind the implementation of Umbraco.Community.Examine.OpenXml.

## Based on UmbracoExamine.PDF

This package follows the architecture and patterns established by [UmbracoExamine.PDF](https://github.com/umbraco/UmbracoExamine.PDF), adapted for OpenXml documents instead of PDF files. Key patterns inherited:

- Custom `LuceneIndex` subclass (`OpenXmlLuceneIndex`) with `IIndexDiagnostics`
- `IndexPopulator` for full index rebuilds with paged media queries
- `MediaCacheRefresherNotification` handler for real-time index sync
- `ValueSetValidator` for recycle bin filtering and parent ID validation
- `IComposer`-based auto-registration (no manual startup config needed)

## Multi-targeting Strategy

### Why net8.0 / net9.0 / net10.0?

Umbraco versions map to specific .NET versions:
- Umbraco 13.x requires .NET 8
- Umbraco 16.x requires .NET 9
- Umbraco 17.x requires .NET 10

A single NuGet package with conditional `Umbraco.Cms.Web.Common` references per TFM allows one package to support all three versions. NuGet automatically selects the correct TFM at install time.

### Why not Umbraco 14 or 15?

Umbraco 14 (.NET 8) and 15 (.NET 9) share TFMs with 13 and 16 respectively. Since only one set of dependencies can exist per TFM in a NuGet package, we target the lowest minor version (13.0.0, 16.0.0, 17.0.0) so consumers on any patch version within that major can use the package. Umbraco 14 users on .NET 8 would get the Umbraco 13 build, which should be compatible.

### Why only Umbraco.Cms.Web.Common?

The original template included `Umbraco.Cms.Api.Common` and `Umbraco.Cms.Api.Management`, but grep confirmed no code references types from those packages. Removing them eliminated a version conflict where the API packages for Umbraco 14+ couldn't coexist with Umbraco 13 types in the same restore.

## Text Extraction Approach

### Word Documents — Paragraph.InnerText vs OpenXmlReader

**Problem**: Word splits text across multiple `<w:r>` (run) elements based on editing history and formatting. A word like "Test" can become `<w:t>T</w:t>` and `<w:t>est</w:t>` in separate runs.

**Original approach**: `OpenXmlReader.GetText()` reads each `<w:t>` element individually. Adding a space between fragments produced "T est" instead of "Test".

**Solution**: `Paragraph.InnerText` correctly concatenates all text within a paragraph's descendant `<w:t>` elements. Spaces are added between paragraphs (not between runs within a paragraph).

This approach also naturally captures text inside tables, text boxes, and shapes since they contain `<w:p>` elements as descendants.

### Word — Additional Content Sources

Headers, footers, footnotes, and endnotes live in separate `OpenXmlPart` objects, not in `MainDocumentPart.Document.Body`. Each is iterated separately using the `AppendParagraphs` helper.

### PowerPoint — Drawing.Paragraph

PowerPoint uses `DocumentFormat.OpenXml.Drawing.Paragraph` (not `Wordprocessing.Paragraph`). The same `InnerText` approach works. Speaker notes are in `SlidePart.NotesSlidePart.NotesSlide`.

### Spreadsheet — Shared String Table

Excel stores deduplicated text in a `SharedStringTablePart`. Cells reference strings by index. The shared string table is materialized to a `List<SharedStringItem>` once per extraction for O(1) lookup instead of O(n) `ElementAt()` per cell. `int.TryParse` with bounds checking prevents crashes on malformed indices.

## Resource Management

All three extractors wrap `WordprocessingDocument.Open()`, `PresentationDocument.Open()`, and `SpreadsheetDocument.Open()` in `using` statements. These implement `IDisposable` and hold file handles — without disposal, indexing many documents under load would leak resources.

## Index Lifecycle

| Event | Handler | Action |
|---|---|---|
| Media created/updated | `OpenXmlCacheNotificationHandler` (RefreshNode) | `AddToIndex` (filters by extension) |
| Media deleted | `OpenXmlCacheNotificationHandler` (Remove) | `RemoveFromIndex` |
| Media trashed | `OpenXmlCacheNotificationHandler` (RefreshNode, Trashed=true) | `RemoveFromIndex` |
| Media moved (branch) | `OpenXmlCacheNotificationHandler` (RefreshBranch) | Processes all descendants |
| Media not found | `OpenXmlCacheNotificationHandler` (null from GetById) | `RemoveFromIndex` |
| Full rebuild | `OpenXmlIndexPopulator.PopulateIndexes` | Pages all media, filters by extension |
| Recycle bin (rebuild) | `OpenXmlValueSetValidator` | Marks as Filtered → deleted from index |

## Search Implementation

The search view uses `GroupedOr` across `fileTextContent` and `nodeName` fields, allowing users to search by document content or filename. The `OpenXmlCategory` scope ensures only OpenXml documents are returned, not other Examine results.

## Umbraco 17 — DevelopmentMode.Backoffice

Umbraco 17 requires the `Umbraco.Cms.DevelopmentMode.Backoffice` NuGet package when using `ModelsMode: InMemoryAuto`. Without it, runtime Razor compilation fails and all front-end views return 404 with a misleading "No physical template file was found" log message.
