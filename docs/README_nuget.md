# Umbraco.Community.Examine.OpenXml

[![Downloads](https://img.shields.io/nuget/dt/Umbraco.Community.Examine.OpenXml?color=cc9900)](https://www.nuget.org/packages/Umbraco.Community.Examine.OpenXml/)
[![NuGet](https://img.shields.io/nuget/vpre/Umbraco.Community.Examine.OpenXml?color=0273B3)](https://www.nuget.org/packages/Umbraco.Community.Examine.OpenXml)
[![GitHub license](https://img.shields.io/github/license/justin-nevitech/Umbraco.Community.Examine.OpenXml?color=8AB803)](https://github.com/justin-nevitech/Umbraco.Community.Examine.OpenXml/blob/main/LICENSE)

An Umbraco package that extracts text content from OpenXml documents (`.docx`, `.pptx`, `.xlsx`) uploaded to the media library and indexes it using [Examine](https://github.com/Shazwazza/Examine), making the content of your Office documents fully searchable.

## How It Works

When OpenXml media files are uploaded or updated in Umbraco, this package:

1. Detects the file type (Word, PowerPoint, or Excel)
2. Extracts the text content using the [DocumentFormat.OpenXml](https://github.com/dotnet/Open-XML-SDK) SDK
3. Indexes the extracted text into a dedicated Lucene-based Examine index called `OpenXmlIndex`

The index is automatically kept in sync when media items are created, updated, or deleted. It can also be rebuilt from the Examine Management dashboard in the Umbraco backoffice.

## Supported File Types

| Extension | Type |
|---|---|
| `.docx` | Word documents |
| `.pptx` | PowerPoint presentations |
| `.xlsx` | Excel spreadsheets |

## Supported Umbraco Versions

| Umbraco | .NET | Status |
|---|---|---|
| 13.x | .NET 8 | Supported |
| 14.x | .NET 8 | Not supported (EOL) |
| 15.x | .NET 9 | Not supported (EOL) |
| 16.x | .NET 9 | Supported |
| 17.x | .NET 10 | Supported |

## Installation

Add the package to an existing Umbraco website from NuGet:

```bash
dotnet add package Umbraco.Community.Examine.OpenXml
```

No additional startup configuration is needed. The package auto-registers via an Umbraco composer and the `OpenXmlIndex` will be available immediately after restarting the site.

## Searching the Index

You can query the `OpenXmlIndex` from any Razor view or controller using Examine's `IExamineManager`:

```csharp
@using Examine
@using Examine.Search
@using Umbraco.Community.Examine.OpenXml

@inject IExamineManager ExamineManager

@{
    var searchQuery = Context.Request.Query["q"].ToString();
}

@if (!string.IsNullOrWhiteSpace(searchQuery))
{
    if (ExamineManager.TryGetIndex(OpenXmlIndexConstants.OpenXmlIndexName, out var index))
    {
        var searcher = index.Searcher;
        var query = searcher.CreateQuery(OpenXmlIndexConstants.OpenXmlCategory)
            .GroupedOr(new[] { OpenXmlIndexConstants.OpenXmlContentFieldName, "nodeName" }, searchQuery);

        var results = query.Execute();

        <p>Found @results.TotalItemCount result(s) for "@searchQuery"</p>

        foreach (var result in results)
        {
            var name = result.Values.ContainsKey("nodeName")
                ? result.Values["nodeName"]
                : "Unknown";

            <div>
                <h2>@name</h2>
                <p>Score: @result.Score.ToString("F2")</p>
            </div>
        }
    }
}
```

### Available Constants

The `OpenXmlIndexConstants` class provides strongly-typed constants for querying the index:

| Constant | Value | Usage |
|---|---|---|
| `OpenXmlIndexName` | `"OpenXmlIndex"` | Index name for `IExamineManager.TryGetIndex()` |
| `OpenXmlContentFieldName` | `"fileTextContent"` | Field name containing the extracted document text |
| `OpenXmlCategory` | `"openxml"` | Category to scope queries to OpenXml documents |

## Limits

To protect against malicious or oversized documents, the following limits are applied during text extraction:

| Limit | Value | Description |
|---|---|---|
| Max file size | 100 MB | Files exceeding this size are skipped entirely |
| Max extracted content | 10 MB | Text extraction stops once this limit is reached |
| Max characters per part | 10,000,000 | Limits characters per OpenXml document part to prevent decompression bombs |
| Max shared strings (Excel) | 1,000,000 | Caps the number of shared string entries loaded from `.xlsx` files |

Documents that exceed these limits are logged as warnings and excluded from the index. These values are defined in the `OpenXmlIndexConstants` class.

## Acknowledgments

This package is based on [UmbracoExamine.PDF](https://github.com/umbraco/UmbracoExamine.PDF) by the Umbraco team. Thank you to the Umbraco HQ developers and contributors for providing the foundation and patterns that made this package possible.

[Files and folders icons created by kawalanicon - Flaticon](https://www.flaticon.com/free-icons/files-and-folders)

For full documentation and source code, visit the [GitHub repository](https://github.com/justin-nevitech/Umbraco.Community.Examine.OpenXml).