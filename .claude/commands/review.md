Review ALL source files in `src/Umbraco.Community.Examine.OpenXml/` for:

## Resource Management
- All `IDisposable` objects (OpenXml documents, streams, readers) wrapped in `using` statements
- No file handle or memory leaks under load

## Null Safety
- No unguarded `.` access on nullable references
- `TryParse` instead of `Parse` for external input
- Bounds checking on collection access (e.g. shared string table indices)

## Exception Handling
- No generic `throw new Exception()` — use specific types
- Notification handlers must not throw (log and return instead)
- Extraction failures must not crash the indexing pipeline

## Examine Index Consistency
- Media create/update/delete/trash all handled in `OpenXmlCacheNotificationHandler`
- `OpenXmlValueSetValidator` filters recycle bin items during rebuild
- `OpenXmlIndexPopulator` correctly pages through all media

## Text Extraction Completeness
- Word: body, headers, footers, footnotes, endnotes extracted via `Paragraph.InnerText`
- PowerPoint: slide content + speaker notes extracted via `Drawing.Paragraph.InnerText`
- Spreadsheet: all cells across all worksheets, shared strings resolved correctly
- No split-run issues (Word runs within same paragraph must concatenate without spaces)

## Multi-major support
- New sources live in the shared folder `src/Umbraco.Community.Examine.OpenXml/`, not in a `.vNN` wrapper project
- Each wrapper project's Umbraco range floor is the lowest release of its major (13.0.0, 17.0.0, 18.0.0)
- No `#if` directives needed — code should be identical across every Umbraco major
- Build every variant: `dotnet build src/Umbraco.Community.Examine.OpenXml.slnx`

Report ALL findings with file paths and line numbers. Flag severity as Critical, High, Medium, or Low.
