using Moq;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public class OpenXmlTextExtractorFactoryTests
{
    private readonly Mock<IWordProcessingDocumentTextExtractor> _wordExtractor = new();
    private readonly Mock<IPresentationDocumentTextExtractor> _presentationExtractor = new();
    private readonly Mock<ISpreadsheetDocumentTextExtractor> _spreadsheetExtractor = new();
    private readonly OpenXmlTextExtractorFactory _factory;

    public OpenXmlTextExtractorFactoryTests()
    {
        _factory = new OpenXmlTextExtractorFactory(
            _wordExtractor.Object,
            _presentationExtractor.Object,
            _spreadsheetExtractor.Object);
    }

    [Fact]
    public void GetOpenXmlTextExtractor_Docx_ReturnsWordExtractor()
    {
        var result = _factory.GetOpenXmlTextExtractor("docx");

        Assert.Same(_wordExtractor.Object, result);
    }

    [Fact]
    public void GetOpenXmlTextExtractor_Pptx_ReturnsPresentationExtractor()
    {
        var result = _factory.GetOpenXmlTextExtractor("pptx");

        Assert.Same(_presentationExtractor.Object, result);
    }

    [Fact]
    public void GetOpenXmlTextExtractor_Xlsx_ReturnsSpreadsheetExtractor()
    {
        var result = _factory.GetOpenXmlTextExtractor("xlsx");

        Assert.Same(_spreadsheetExtractor.Object, result);
    }

    [Fact]
    public void GetOpenXmlTextExtractor_UnsupportedExtension_ThrowsNotSupportedException()
    {
        Assert.Throws<NotSupportedException>(() =>
            _factory.GetOpenXmlTextExtractor("pdf"));
    }

    [Fact]
    public void GetOpenXmlTextExtractor_EmptyExtension_ThrowsNotSupportedException()
    {
        Assert.Throws<NotSupportedException>(() =>
            _factory.GetOpenXmlTextExtractor(""));
    }

    [Fact]
    public void GetOpenXmlTextExtractor_DocxWithDot_ThrowsNotSupportedException()
    {
        // The factory expects extensions without a leading dot
        Assert.Throws<NotSupportedException>(() =>
            _factory.GetOpenXmlTextExtractor(".docx"));
    }

    [Fact]
    public void GetOpenXmlTextExtractor_UpperCaseExtension_ReturnsCorrectExtractor()
    {
        // Dictionary uses case-insensitive comparison, so "DOCX" should match "docx"
        var extractor = _factory.GetOpenXmlTextExtractor("DOCX");
        Assert.IsAssignableFrom<IWordProcessingDocumentTextExtractor>(extractor);
    }

    [Fact]
    public void GetOpenXmlTextExtractor_MixedCaseExtension_ReturnsCorrectExtractor()
    {
        var pptxExtractor = _factory.GetOpenXmlTextExtractor("Pptx");
        Assert.IsAssignableFrom<IPresentationDocumentTextExtractor>(pptxExtractor);

        var xlsxExtractor = _factory.GetOpenXmlTextExtractor("xLsX");
        Assert.IsAssignableFrom<ISpreadsheetDocumentTextExtractor>(xlsxExtractor);
    }

    [Fact]
    public void GetOpenXmlTextExtractor_NullExtension_ThrowsException()
    {
        Assert.ThrowsAny<Exception>(() =>
            _factory.GetOpenXmlTextExtractor(null!));
    }
}
