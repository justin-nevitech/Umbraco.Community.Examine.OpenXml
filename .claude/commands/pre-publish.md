Run the full pre-publish checklist for the Umbraco.Community.Examine.OpenXml package.

## 1. Build Solution
```
dotnet build src/Umbraco.Community.Examine.OpenXml.slnx -c Release
```
- Must be 0 errors
- Report any code warnings (ignore NuGet vulnerability warnings from Umbraco dependencies)

## 2. Run Tests
```
dotnet test src/Umbraco.Community.Examine.OpenXml.slnx -c Release
```
- All tests must pass, on every supported Umbraco major (Tests.v13, Tests.v17, Tests.v18)

## 3. Pack and Inspect
Releases are version-aligned — one nupkg per Umbraco major, all sharing one PackageId. Pack each
variant with a version whose major matches it:
```
dotnet pack src/Umbraco.Community.Examine.OpenXml.v13/Umbraco.Community.Examine.OpenXml.v13.csproj -c Release -p:Version=13.0.0 -o /tmp/nupkg-check
dotnet pack src/Umbraco.Community.Examine.OpenXml.v17/Umbraco.Community.Examine.OpenXml.v17.csproj -c Release -p:Version=17.0.0 -o /tmp/nupkg-check
dotnet pack src/Umbraco.Community.Examine.OpenXml.v18/Umbraco.Community.Examine.OpenXml.v18.csproj -c Release -p:Version=18.0.0 -o /tmp/nupkg-check
```
Verify each nupkg contains:
- A single `lib/<tfm>/Umbraco.Community.Examine.OpenXml.dll` — `net8.0` for v13, `net10.0` for v17 and v18
- `README_nuget.md`
- `icon.png`
- `LICENSE`
- Correct nuspec metadata (ID, title, description, authors, license, tags, repository URL) — identical across all three
- The right Umbraco range in the dependency group: `[13.0.0, 14.0.0)`, `[17.0.0, 18.0.0)`, `[18.0.0, 19.0.0)`

## 4. Verify Documentation
Check these files are up to date:
- `.github/README.md` — Supported versions table, install commands, code sample, author, acknowledgments
- `docs/README_nuget.md` — Same as above, tailored for NuGet
- `docs/BUILDING.md` — Project layout, ranges, packaging model and test-site table match reality
- `umbraco-marketplace.json` — Category, description, tags, icon URL, title
- `CLAUDE.md` — Architecture and commands accurate

## 5. Verify CI/CD
- `.github/workflows/release.yml` exists and maps every supported tag major (13, 17, 18) to the right package and test project
- A tag major with no matching project fails the job rather than publishing the wrong variant
- The suite for the major being published runs before the pack step
- Version is injected via `/p:Version=${{github.ref_name}}`
- `NUGET_API_KEY` secret is referenced

## 6. Report
Summarize the results as a checklist with pass/fail for each item.
