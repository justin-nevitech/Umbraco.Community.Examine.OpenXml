using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public class SpreadsheetDocumentTextExtractorTests
{
    private readonly SpreadsheetDocumentTextExtractor _extractor = new();

    [Fact]
    public void GetText_WithRealXlsx_ExtractsText()
    {
        using var stream = TestHelper.GetTestFileStream("test.xlsx");

        var result = _extractor.GetText(stream);

        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void GetText_WithRealXlsx_ContainsCellContent()
    {
        using var stream = TestHelper.GetTestFileStream("test.xlsx");

        var result = _extractor.GetText(stream);

        // Verify that text from cells is extracted
        Assert.NotNull(result);
        Assert.True(result.Length > 0, "Should extract text from spreadsheet cells");
    }

    [Fact]
    public void GetText_StreamIsNotDisposedByExtractor()
    {
        using var stream = TestHelper.GetTestFileStream("test.xlsx");

        _extractor.GetText(stream);

        Assert.True(stream.CanRead);
    }

    [Fact]
    public void GetText_ReturnsString_NotNull()
    {
        using var stream = TestHelper.GetTestFileStream("test.xlsx");

        var result = _extractor.GetText(stream);

        Assert.IsType<string>(result);
    }

    [Fact]
    public void GetText_WithEmptySpreadsheet_ReturnsEmptyString()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets());
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void GetText_WithSpreadsheetContainingOnlyNumericValues_ExtractsNumbers()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    new Cell { CellValue = new CellValue("42"), DataType = CellValues.Number },
                    new Cell { CellValue = new CellValue("3.14"), DataType = CellValues.Number })));

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" }));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("42", result);
        Assert.Contains("3.14", result);
    }

    [Fact]
    public void GetText_WithSpreadsheetWithoutSharedStringTable_ExtractsInlineValues()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            // No SharedStringTablePart added
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    new Cell { CellValue = new CellValue("InlineValue1") },
                    new Cell { CellValue = new CellValue("InlineValue2") })));

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" }));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("InlineValue1", result);
        Assert.Contains("InlineValue2", result);
    }

    [Fact]
    public void GetText_WithMultipleWorksheets_ExtractsFromAll()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();

            var worksheetPart1 = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart1.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell { CellValue = new CellValue("FromSheet1") })));

            var worksheetPart2 = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart2.Worksheet = new Worksheet(new SheetData(
                new Row(new Cell { CellValue = new CellValue("FromSheet2") })));

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart1), SheetId = 1, Name = "Sheet1" },
                new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart2), SheetId = 2, Name = "Sheet2" }));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("FromSheet1", result);
        Assert.Contains("FromSheet2", result);
    }
}
