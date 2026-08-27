# Contributing Guidelines

Contributions to this package are most welcome!

## Getting Started

There are three test sites in the solution to make working with this repository easier, one per
supported Umbraco major. Each is configured to do an unattended install — check `appSettings.json`
for the login details.

### Prerequisites

- [.NET 8 SDK](https://dotnet.microsoft.com/download) (for the Umbraco 13 projects)
- [.NET 10 SDK](https://dotnet.microsoft.com/download) (for the Umbraco 17 and 18 projects)

### Running a Test Site

1. Clone the repository
2. Open `src/Umbraco.Community.Examine.OpenXml.slnx` in your IDE
3. Set one of the test sites as the startup project
4. Run it — it will perform an unattended Umbraco install on first run
5. Log in with the credentials from `appSettings.json`, then upload a `.docx`, `.pptx` or `.xlsx`
   to the media library and search for its contents from the `openXmlSearch` page

| Site | Umbraco | URL |
| ---- | ------- | --- |
| `Umbraco.Community.Examine.OpenXml.TestSite.v13` | 13 | `https://localhost:44380` |
| `Umbraco.Community.Examine.OpenXml.TestSite.v17` | 17 | `https://localhost:44379` |
| `Umbraco.Community.Examine.OpenXml.TestSite.v18` | 18 | `https://localhost:44381` |

All three are in the solution, listen on different ports and can run at the same time. Each
references the package project for its own Umbraco major.

> The package supports Umbraco 13, 17 and 18 from one set of sources, via three wrapper package
> projects that compile the same files (`Umbraco.Community.Examine.OpenXml.v13`, `.v17` and
> `.v18`). The sources themselves live in `Umbraco.Community.Examine.OpenXml`, which is a shared
> source folder rather than a project — add new files there and all three variants pick them up
> automatically. See [docs/BUILDING.md](../docs/BUILDING.md) for the multi-major build details.

### Running the Tests

```bash
cd src
dotnet test Umbraco.Community.Examine.OpenXml.slnx        # every supported major
dotnet test Umbraco.Community.Examine.OpenXml.Tests.v13   # Umbraco 13 only
dotnet test Umbraco.Community.Examine.OpenXml.Tests.v17   # Umbraco 17 only
dotnet test Umbraco.Community.Examine.OpenXml.Tests.v18   # Umbraco 18 only
```

The test sources live once in `src/Umbraco.Community.Examine.OpenXml.Tests` (a shared folder, not
a project) and are compiled by all three wrapper test projects, so every test runs against every
Umbraco major. Add new tests there and all three pick them up.

Tests must not require a running Umbraco instance or a Lucene index. Extractor tests use the real
sample documents in `TestFiles`; edge cases build documents in memory with `DocumentFormat.OpenXml`.

## Releasing

Releases are version-aligned — the package major tracks the Umbraco major, and all three variants
publish under one PackageId. Pushing a tag such as `18.0.0` runs the Umbraco 18 test suite, then
packs and publishes the `.v18` project. See [docs/BUILDING.md](../docs/BUILDING.md).
