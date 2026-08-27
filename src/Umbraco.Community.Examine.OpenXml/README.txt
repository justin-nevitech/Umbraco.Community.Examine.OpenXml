== Umbraco.Community.Examine.OpenXml - shared sources ==

This folder is NOT a project. It has no .csproj.

It is the single home for the package's C# sources. Three wrapper projects compile these same
files against a different Umbraco major, and produce the same assembly name and PackageId:

  ..\Umbraco.Community.Examine.OpenXml.v13   net8.0    Umbraco [13.0.0, 14.0.0)   package 13.x
  ..\Umbraco.Community.Examine.OpenXml.v17   net10.0   Umbraco [17.0.0, 18.0.0)   package 17.x
  ..\Umbraco.Community.Examine.OpenXml.v18   net10.0   Umbraco [18.0.0, 19.0.0)   package 18.x

The shared compile items and the package metadata are declared once in
Examine.OpenXml.Shared.props, which each wrapper imports. Add new source files to THIS folder and
all three variants pick them up automatically - never add sources to a wrapper project.

The sources are identical across every major: there are no #if preprocessor directives. The
package only uses Umbraco and Examine APIs that are unchanged from 13 through 18.

== Build ==

  dotnet build ..\Umbraco.Community.Examine.OpenXml.slnx                                    all variants
  dotnet build ..\Umbraco.Community.Examine.OpenXml.v18\Umbraco.Community.Examine.OpenXml.v18.csproj

== Tests ==

The test sources mirror this layout and live in ..\Umbraco.Community.Examine.OpenXml.Tests, which
is also a shared folder rather than a project. Three wrapper test projects (.Tests.v13, .Tests.v17,
.Tests.v18) run the whole suite against each supported major.

  dotnet test ..\Umbraco.Community.Examine.OpenXml.slnx

== Full detail ==

See docs\BUILDING.md in the repository root for the project layout, the version-aligned packaging
model and the release process.
