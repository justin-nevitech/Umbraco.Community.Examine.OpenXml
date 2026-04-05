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
    public void GetText_WithSharedStringIndexOutOfRange_DoesNotThrow()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();

            // Create a shared string table with only 2 entries
            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new SharedStringTable(
                new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("Hello")),
                new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("World")));

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    // Valid index
                    new Cell { CellValue = new CellValue("0"), DataType = CellValues.SharedString },
                    // Out-of-range index — should not throw
                    new Cell { CellValue = new CellValue("99"), DataType = CellValues.SharedString })));

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" }));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("Hello", result);
        // Out-of-range index falls back to CellValue.InnerText ("99")
        Assert.Contains("99", result);
    }

    [Fact]
    public void GetText_WithNonNumericSharedStringReference_DoesNotThrow()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();

            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new SharedStringTable(
                new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("Hello")));

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    // Non-numeric value where int.TryParse will fail
                    new Cell { CellValue = new CellValue("abc"), DataType = CellValues.SharedString })));

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" }));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        // Falls back to CellValue.InnerText
        Assert.Contains("abc", result);
    }

    [Fact]
    public void GetText_WithCellsContainingNullCellValue_DoesNotThrow()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    new Cell(), // No CellValue at all
                    new Cell { CellValue = new CellValue("ValidValue") })));

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" }));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("ValidValue", result);
    }

    [Fact]
    public void GetText_WithEmptyCellValue_SkipsEmptyCells()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    new Cell { CellValue = new CellValue("") },
                    new Cell { CellValue = new CellValue("   ") },
                    new Cell { CellValue = new CellValue("ActualData") })));

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" }));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("ActualData", result);
    }

    [Fact]
    public void GetText_WithSharedStringHavingNullText_DoesNotThrow()
    {
        using var stream = new MemoryStream();
        using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();

            // Shared string table with an entry that has no Text child
            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new SharedStringTable(
                new SharedStringItem(), // No Text element
                new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("ValidShared")));

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData(
                new Row(
                    new Cell { CellValue = new CellValue("0"), DataType = CellValues.SharedString },
                    new Cell { CellValue = new CellValue("1"), DataType = CellValues.SharedString })));

            workbookPart.Workbook = new Workbook(new Sheets(
                new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" }));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("ValidShared", result);
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
