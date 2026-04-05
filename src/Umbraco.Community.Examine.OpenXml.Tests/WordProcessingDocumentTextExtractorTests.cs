using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public class WordProcessingDocumentTextExtractorTests
{
    private readonly WordProcessingDocumentTextExtractor _extractor = new();

    [Fact]
    public void GetText_WithRealDocx_ExtractsText()
    {
        using var stream = TestHelper.GetTestFileStream("test.docx");

        var result = _extractor.GetText(stream);

        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void GetText_WithRealDocx_ContainsExpectedContent()
    {
        using var stream = TestHelper.GetTestFileStream("test.docx");

        var result = _extractor.GetText(stream);

        // The docx has split runs: "T" + "est" + " – this content should get indexed"
        Assert.Contains("Test", result);
        Assert.Contains("this content should get indexed", result);
    }

    [Fact]
    public void GetText_WithRealDocx_ConcatenatesSplitRuns()
    {
        using var stream = TestHelper.GetTestFileStream("test.docx");

        var result = _extractor.GetText(stream);

        // "T" and "est" from separate runs should be concatenated into "Test"
        Assert.Contains("Test", result);
        Assert.DoesNotContain("T est", result);
    }

    [Fact]
    public void GetText_StreamIsNotDisposedByExtractor()
    {
        using var stream = TestHelper.GetTestFileStream("test.docx");

        _extractor.GetText(stream);

        // Stream should still be usable after extraction (not disposed by the extractor)
        Assert.True(stream.CanRead);
    }

    [Fact]
    public void GetText_ReturnsString_NotNull()
    {
        using var stream = TestHelper.GetTestFileStream("test.docx");

        var result = _extractor.GetText(stream);

        Assert.IsType<string>(result);
    }

    [Fact]
    public void GetText_WithEmptyDocument_ReturnsEmptyString()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body());
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void GetText_WithHeaderAndFooter_ExtractsHeaderAndFooterText()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("BodyText")))));

            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new Header(new Paragraph(new Run(new Text("HeaderText"))));

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new Footer(new Paragraph(new Run(new Text("FooterText"))));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("BodyText", result);
        Assert.Contains("HeaderText", result);
        Assert.Contains("FooterText", result);
    }

    [Fact]
    public void GetText_WithFootnoteAndEndnote_ExtractsBoth()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("BodyText")))));

            var footnotesPart = mainPart.AddNewPart<FootnotesPart>();
            footnotesPart.Footnotes = new Footnotes(
                new Footnote(new Paragraph(new Run(new Text("FootnoteText")))) { Id = 1 });

            var endnotesPart = mainPart.AddNewPart<EndnotesPart>();
            endnotesPart.Endnotes = new Endnotes(
                new Endnote(new Paragraph(new Run(new Text("EndnoteText")))) { Id = 1 });
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("BodyText", result);
        Assert.Contains("FootnoteText", result);
        Assert.Contains("EndnoteText", result);
    }

    [Fact]
    public void GetText_WithMultipleSplitRuns_ConcatenatesCorrectly()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            // Single paragraph with 4 runs: "Hel" + "lo" + " " + "World"
            mainPart.Document = new Document(new Body(
                new Paragraph(
                    new Run(new Text("Hel") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new Text("lo") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new Text("World")))));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("Hello World", result);
    }

    [Fact]
    public void GetText_WithWhitespaceOnlyParagraphs_DoesNotAddExtraContent()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(new Body(
                new Paragraph(new Run(new Text("   "))),
                new Paragraph(new Run(new Text("RealContent"))),
                new Paragraph()));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("RealContent", result);
    }

    [Fact]
    public void GetText_WithDocumentContainingTable_ExtractsTableContent()
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            var table = new Table(
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("Cell1")))),
                    new TableCell(new Paragraph(new Run(new Text("Cell2"))))),
                new TableRow(
                    new TableCell(new Paragraph(new Run(new Text("Cell3")))),
                    new TableCell(new Paragraph(new Run(new Text("Cell4"))))));
            mainPart.Document = new Document(new Body(table));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("Cell1", result);
        Assert.Contains("Cell2", result);
        Assert.Contains("Cell3", result);
        Assert.Contains("Cell4", result);
    }
}
