# Discovery

Findings, issues encountered, and lessons learned during development.

## Text Extraction Issues

### Split Runs in Word Documents

**Discovery**: Word documents frequently split a single word across multiple `<w:r>` elements due to revision tracking. The sample `test.docx` had "Test" stored as:
```xml
<w:r><w:t>T</w:t></w:r>
<w:r w:rsidR="007667B6"><w:t>est</w:t></w:r>
```

**Impact**: The original `OpenXmlReader`-based extraction added spaces between every text fragment, producing `"T est"` in the Lucene index. This made the first word unsearchable.

**Resolution**: Switched to `Paragraph.InnerText` which correctly concatenates all runs within a paragraph.

### AppendLine Injecting \r\n Into Indexed Content

**Discovery**: The original code used `StringBuilder.AppendLine()` between text fragments. The `\r\n` characters were embedded into the indexed content, breaking Lucene tokenization. The indexed value for "Test" would include `\r\n` making it unmatchable as a search token.

**Resolution**: Replaced all `AppendLine` with `Append` + space separator.

### Empty Text From test.docx

**Discovery**: Initial testing showed the sample `test.docx` appeared to have empty text content when extracted via PowerShell. This was because the text was split across runs and the extraction tool was not handling it correctly — the document did contain text ("Test – this content should get indexed").

## Umbraco Version Compatibility

### Umbraco 17 Front-End 404 Error

**Discovery**: After setting up the test site with the Clean starter kit on Umbraco 17, ALL front-end pages returned 404 with the log message "No physical template file was found for template home" — even though the .cshtml files existed on disk.

**Root cause**: The Umbraco log revealed a fatal error: `InMemoryAuto requires the Umbraco.Cms.DevelopmentMode.Backoffice package to be installed`. Without this package, runtime Razor compilation doesn't work, causing all views to be unresolvable.

**Resolution**: Added `Umbraco.Cms.DevelopmentMode.Backoffice` package to the v17 test site.

### API Package Version Conflict

**Discovery**: When adding multi-targeting, `dotnet restore` failed with `NU1107: Version conflict detected for Umbraco.Cms.Core`. The `Umbraco.Cms.Api.Management` package for v13 didn't exist (the Management API was introduced in Umbraco 14), causing NuGet to resolve to v14 which conflicted with the v13 Core package.

**Resolution**: Removed `Umbraco.Cms.Api.Common` and `Umbraco.Cms.Api.Management` references entirely — grep confirmed no source code uses types from these packages.

### Clean Starter Kit Version Mapping

**Discovery**: The Clean NuGet package versions don't follow Umbraco versioning:
- Clean 4.2.2 → Umbraco 13 (net8.0)
- Clean 5.x → Umbraco 15 (net9.0) — NOT for Umbraco 16
- Clean 6.0.1 → Umbraco 16 (net9.0)
- Clean 7.0.5 → Umbraco 17 (net10.0)

Verified by inspecting each package's `.nuspec` dependency on `Umbraco.Cms.Web.Website`.

### Smidge Requirement for Umbraco 13

**Discovery**: The v13 test site failed with `The name 'SmidgeHelper' does not exist in the current context` on the Clean starter kit's `master.cshtml`. Umbraco 13 has a transitive dependency on Smidge but the helpers require explicit registration.

**Resolution**: Added `using Smidge;`, `builder.Services.AddSmidge()`, and `app.UseSmidge()` to the v13 `Program.cs`. No explicit NuGet package reference needed since Umbraco already provides Smidge transitively.

### uSync Folder Naming

**Discovery**: uSync v13.x stores files under `uSync/v9/` while uSync v16.x and v17.x use `uSync/v17/`. The folder name relates to the uSync format version, not the Umbraco version.

## Resource Leaks

**Discovery**: Production review found that all three OpenXml document objects (`WordprocessingDocument`, `PresentationDocument`, `SpreadsheetDocument`) implement `IDisposable` but were not being disposed. Under load with many media uploads, this would leak file handles and memory.

**Resolution**: Wrapped all three in `using` statements.

## Spreadsheet Parsing Edge Cases

### Int32.Parse on Empty String

**Discovery**: `Int32.Parse(cell?.CellValue?.InnerText ?? String.Empty)` would throw `FormatException` if `CellValue` was null, since `String.Empty` can't be parsed as an integer.

**Resolution**: Replaced with `int.TryParse` with bounds checking against the shared string list count.

### ElementAt Performance

**Discovery**: `Elements<SharedStringItem>().ElementAt(index)` is O(n) per cell lookup. For a spreadsheet with 10,000 shared strings and 50,000 cells, this means 500 million element traversals.

**Resolution**: Materialized the shared string table to a `List<T>` once, making each lookup O(1) via index.

## Notification Handler Issues

### NotSupportedException Crash Risk

**Discovery**: The `OpenXmlCacheNotificationHandler` threw `NotSupportedException` for any unexpected `MessageType`. If Umbraco introduced a new message type in a future version, this would crash the entire notification pipeline.

**Resolution**: Replaced `throw` with `LogWarning` + `return`.

### RefreshAll Silently Ignored

**Discovery**: The `RefreshAll` change type was silently ignored with a dismissive comment. While Examine doesn't support RefreshAll as an incremental operation, the silent handling made debugging difficult.

**Resolution**: Added `LogDebug` message explaining that a full index rebuild should be triggered from the Examine Management dashboard.

## NuGet Package Issues

### Duplicated Package ID

**Discovery**: The project template generated `Umbraco.Community.Umbraco.Community.Examine.OpenXml` as the PackageId — duplicating the `Umbraco.Community.` prefix. This affected badge URLs, install commands, and the NuGet listing.

**Resolution**: Corrected to `Umbraco.Community.Examine.OpenXml` across .csproj, both READMEs, and marketplace JSON.

### Missing LICENSE in nupkg

**Discovery**: Comparing with the Umbraco.Community.AI.LogAnalyser package revealed the LICENSE file was not being included in the NuGet package, even though `PackageLicenseExpression` was set.

**Resolution**: Added LICENSE as a `<None>` pack item in the .csproj.

## Test Coverage

83 unit tests cover all public functionality without requiring an Umbraco instance or Lucene index. Real `.docx`, `.pptx`, and `.xlsx` files from the `samples/` folder are used for extractor tests. In-memory OpenXml documents created via `DocumentFormat.OpenXml` cover edge cases (empty documents, tables, numeric cells, multiple worksheets, speaker notes).
