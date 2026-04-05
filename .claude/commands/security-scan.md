Perform a security review of ALL source files in `src/Umbraco.Community.Examine.OpenXml/`.

## Dependency Review
Run `dotnet list src/Umbraco.Community.Examine.OpenXml/Umbraco.Community.Examine.OpenXml.csproj package --vulnerable` for each TFM and report any known CVEs.

## Code Review

### Input Validation
- File paths from Umbraco media are used to open files — check for path traversal risks
- File extensions are parsed — check for injection via crafted extensions
- Shared string indices from Excel files are parsed — check for out-of-bounds access

### Resource Safety
- All `IDisposable` objects must be in `using` statements
- `OpenXmlReader`, `StreamReader`, and OpenXml document objects checked
- No unbounded memory allocation from malicious documents

### Exception Information Disclosure
- Error messages and log entries must not expose internal paths or stack traces to end users
- Exceptions from document parsing must be caught and logged, not propagated to HTTP responses

### Denial of Service
- Check for unbounded loops (e.g. `GetPagedDescendants` pagination)
- Check for excessive memory usage from large documents (very large spreadsheets, documents with millions of paragraphs)
- Shared string table is materialized to a List — check impact with very large tables

### OWASP Considerations
- No user input is rendered unescaped in the search view (XSS)
- Search query parameter is used in Examine query — check for Lucene query injection
- No SQL or command injection vectors

Report findings with severity (Critical/High/Medium/Low), file path, line number, and recommended fix.
