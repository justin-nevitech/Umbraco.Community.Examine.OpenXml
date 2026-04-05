Run the full pre-publish checklist for the Umbraco.Community.Examine.OpenXml package.

## 1. Build Solution
```
dotnet build src/Umbraco.Community.Examine.OpenXml.slnx -c Release
```
- Must be 0 errors
- Report any code warnings (ignore NuGet vulnerability warnings from Umbraco dependencies)

## 2. Run Tests
```
dotnet test src/Umbraco.Community.Examine.OpenXml.Tests/Umbraco.Community.Examine.OpenXml.Tests.csproj
```
- All tests must pass

## 3. Pack and Inspect
```
dotnet pack src/Umbraco.Community.Examine.OpenXml/Umbraco.Community.Examine.OpenXml.csproj -c Release -o /tmp/nupkg-check
```
Verify the nupkg contains:
- `lib/net8.0/`, `lib/net9.0/`, `lib/net10.0/` DLLs
- `README_nuget.md`
- `icon.png`
- `LICENSE`
- Correct nuspec metadata (ID, title, description, authors, license, tags, repository URL)
- Correct per-TFM dependency groups

## 4. Verify Documentation
Check these files are up to date:
- `.github/README.md` — Supported versions table, code sample, acknowledgments
- `docs/README_nuget.md` — Same as above, tailored for NuGet
- `umbraco-marketplace.json` — Category, description, tags, icon URL, title
- `CLAUDE.md` — Architecture and commands accurate

## 5. Verify CI/CD
- `.github/workflows/release.yml` exists and references correct .csproj path
- Version is injected via `/p:Version=${{github.ref_name}}`
- `NUGET_API_KEY` secret is referenced

## 6. Report
Summarize the results as a checklist with pass/fail for each item.
