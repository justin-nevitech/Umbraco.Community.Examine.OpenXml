# Building & multi-major (Umbraco 13 / 17 / 18) support

This package supports Umbraco **13**, **17** and **18** from a single set of sources. Umbraco 17
and 18 both run on `net10.0`, so plain TFM-based multi-targeting cannot express the difference
between them. Instead there are **three package projects that compile the same source files**
against different Umbraco package ranges.

## Layout

| Project | TFM | Umbraco.Cms.Web.Common | Produces |
| ------- | --- | ---------------------- | -------- |
| `src/Umbraco.Community.Examine.OpenXml.v13` | `net8.0`  | `[13.0.0, 14.0.0)` | package `13.x` |
| `src/Umbraco.Community.Examine.OpenXml.v17` | `net10.0` | `[17.0.0, 18.0.0)` | package `17.x` |
| `src/Umbraco.Community.Examine.OpenXml.v18` | `net10.0` | `[18.0.0, 19.0.0)` | package `18.x` |

**`src/Umbraco.Community.Examine.OpenXml` is not a project — it is the shared source folder.**
It holds all the C# sources, but no `.csproj`.

The three wrapper projects above own no sources of their own. Each imports
`Examine.OpenXml.Shared.props` from that folder, which declares the shared compile items and the
package metadata using `$(MSBuildThisFileDirectory)`-anchored paths. Add new files to the shared
folder and all three variants pick them up automatically. All three produce the same assembly
name and the same `PackageId`; the *only* differences are the target framework and the Umbraco
range above.

The sources are **identical** across every major — there are no `#if` preprocessor directives.
The package only uses Umbraco and Examine APIs that are unchanged from 13 through 18.

Each floor is the lowest release of its major (`13.0.0`, `17.0.0`, `18.0.0`) so that every patch
release of that major is a valid host.

Build/pack each variant:

```bash
# Umbraco 13
dotnet build src/Umbraco.Community.Examine.OpenXml.v13/Umbraco.Community.Examine.OpenXml.v13.csproj -c Release
dotnet pack  src/Umbraco.Community.Examine.OpenXml.v13/Umbraco.Community.Examine.OpenXml.v13.csproj -c Release /p:Version=13.0.0

# Umbraco 17
dotnet build src/Umbraco.Community.Examine.OpenXml.v17/Umbraco.Community.Examine.OpenXml.v17.csproj -c Release
dotnet pack  src/Umbraco.Community.Examine.OpenXml.v17/Umbraco.Community.Examine.OpenXml.v17.csproj -c Release /p:Version=17.0.0

# Umbraco 18
dotnet build src/Umbraco.Community.Examine.OpenXml.v18/Umbraco.Community.Examine.OpenXml.v18.csproj -c Release
dotnet pack  src/Umbraco.Community.Examine.OpenXml.v18/Umbraco.Community.Examine.OpenXml.v18.csproj -c Release /p:Version=18.0.0
```

The version ranges flow into the packed `.nuspec` dependency nodes, so a `17.x` release depends
on Umbraco `[17.0.0, 18.0.0)` and an `18.x` release on `[18.0.0, 19.0.0)`.

## Why three projects instead of one multi-targeted project

The original layout was a single project with `<TargetFrameworks>net8.0;net9.0;net10.0</TargetFrameworks>`
and a conditional `PackageReference` per TFM. That works only while each supported Umbraco major
sits on its own .NET version. Umbraco 17 and 18 both target `net10.0`, so a TFM list has no way
to say "net10.0 against Umbraco 17" *and* "net10.0 against Umbraco 18" in the same project.

A single project could only express the difference with an MSBuild property such as
`-p:UmbracoMajor=18`, and that fails for the development workflow: **NuGet restore evaluates each
project once with its default properties**, ignoring `AdditionalProperties` metadata on a
`ProjectReference`. One solution could therefore never restore the shared project at both majors
simultaneously, and the v17 and v18 test sites could not coexist in the solution (`NU1107`).

Three project files give each major its own restore, so the whole solution — all three package
variants and all three test sites — builds together and every site can run side by side.

> A package with no version-specific code *can* avoid the split for two adjacent majors by
> declaring one wide range such as `[17.0.0, 19.0.0)`. That is deliberately not done here: each
> major gets its own release line, so a future Umbraco 19 or a 17-only fix can be shipped
> without disturbing the others.

## Packaging model: version-aligned, one PackageId

There is **one** PackageId (`Umbraco.Community.Examine.OpenXml`). The package major tracks the
Umbraco major:

- Install on Umbraco 13 → `13.x` of this package.
- Install on Umbraco 17 → `17.x` of this package.
- Install on Umbraco 18 → `18.x` of this package.

This means the package **does not follow semantic versioning at the major level**: the major
communicates the target Umbraco major, not a breaking change in this package. The three lines are
parallel, not successive — `18.0.0` is not "newer work" than `13.4.0`, it is the same code built
against a different Umbraco. Minor and patch keep their semantic meaning within a line.

A consequence worth repeating in the READMEs: because all three lines share one PackageId, NuGet
resolves a bare `dotnet add package` to the highest version, which is wrong for anyone not on the
newest Umbraco major. Consumers must pin the major.

The release workflow (`.github/workflows/release.yml`) selects the project to pack from the
leading major of the pushed tag, so tagging `17.2.3` packs the `.v17` project and `18.0.0` packs
the `.v18` project. A tag whose major is not 13, 17 or 18 fails the build rather than publishing
the wrong variant.

The previously published `1.x` releases (a single package multi-targeting `net8.0`/`net9.0`/`net10.0`)
remain on NuGet and are unaffected.

## Tests

The test suite mirrors the package layout. `src/Umbraco.Community.Examine.OpenXml.Tests` is a
**shared source folder, not a project**; three wrapper projects compile the same test files
against their respective package variant:

| Project | TFM | Compiles against |
| ------- | --- | ---------------- |
| `src/Umbraco.Community.Examine.OpenXml.Tests.v13` | `net8.0`  | `Umbraco.Community.Examine.OpenXml.v13` |
| `src/Umbraco.Community.Examine.OpenXml.Tests.v17` | `net10.0` | `Umbraco.Community.Examine.OpenXml.v17` |
| `src/Umbraco.Community.Examine.OpenXml.Tests.v18` | `net10.0` | `Umbraco.Community.Examine.OpenXml.v18` |

All 119 tests therefore run on **all three** majors, against the same Umbraco assemblies the
shipped package is compiled against.

```bash
dotnet test src/Umbraco.Community.Examine.OpenXml.slnx          # all three majors
dotnet test src/Umbraco.Community.Examine.OpenXml.Tests.v13     # Umbraco 13 only
dotnet test src/Umbraco.Community.Examine.OpenXml.Tests.v17     # Umbraco 17 only
dotnet test src/Umbraco.Community.Examine.OpenXml.Tests.v18     # Umbraco 18 only
```

Add new tests to the shared folder; all three wrappers pick them up automatically. The sample
`.docx`/`.pptx`/`.xlsx` documents in `TestFiles` are linked from the shared folder by
`Examine.OpenXml.Tests.Shared.props` and copied next to each test assembly.

One construction detail is major-dependent and is handled in `TestHelper.CreateMediaFileManager`:
from Umbraco 18 the public `MediaFileManager` constructor resolves `Lazy<ICoreScopeProvider>` from
`StaticServiceProvider`, which Umbraco populates during boot. Unit tests never boot Umbraco, so
the helper seeds that locator with a stub. On Umbraco 13 and 17 the constructor does not touch the
locator and the assignment is inert, so the one helper is valid on every major.

The suite builds clean: the only warnings in a full build are NuGet vulnerability advisories
(`NU1902`/`NU1903`) for packages that arrive transitively through Umbraco's own dependency tree.

The release workflow runs the suite for the major being published before it packs, so a failing
test blocks the push to NuGet.

## Test sites

All three test sites are in the solution and can run at the same time:

| Site | Umbraco | URL | References |
| ---- | ------- | --- | ---------- |
| `TestSite.v13` | 13 (latest 13.x) | `https://localhost:44380` | `Umbraco.Community.Examine.OpenXml.v13` |
| `TestSite.v17` | 17 (latest 17.x) | `https://localhost:44379` | `Umbraco.Community.Examine.OpenXml.v17` |
| `TestSite.v18` | 18 (latest 18.x) | `https://localhost:44381` | `Umbraco.Community.Examine.OpenXml.v18` |

```bash
dotnet run --project src/Umbraco.Community.Examine.OpenXml.TestSite.v13
dotnet run --project src/Umbraco.Community.Examine.OpenXml.TestSite.v17
dotnet run --project src/Umbraco.Community.Examine.OpenXml.TestSite.v18
```

Unlike the package projects, the test sites track the **newest** release of their Umbraco major,
so each one exercises the package against the latest CMS it will actually be installed on.

All use the Clean starter kit, uSync for content import, the same unattended-install credentials
(admin@example.com / 1234567890) and a local SQLite database, and listen on different ports so all
three can run at once. The v13 site uses the uSync folder `uSync/v9`, v17 uses `uSync/v17` and v18
uses `uSync/v18`.
