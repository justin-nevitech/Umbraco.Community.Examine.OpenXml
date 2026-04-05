namespace Umbraco.Community.Examine.OpenXml.Tests;

public class OpenXmlIndexConstantsTests
{
    [Fact]
    public void OpenXmlIndexName_HasExpectedValue()
    {
        Assert.Equal("OpenXmlIndex", OpenXmlIndexConstants.OpenXmlIndexName);
    }

    [Fact]
    public void OpenXmlContentFieldName_HasExpectedValue()
    {
        Assert.Equal("fileTextContent", OpenXmlIndexConstants.OpenXmlContentFieldName);
    }

    [Fact]
    public void UmbracoMediaExtensionPropertyAlias_HasExpectedValue()
    {
        Assert.Equal("umbracoExtension", OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias);
    }

    [Fact]
    public void WordProcessingDocumentFileExtension_HasExpectedValue()
    {
        Assert.Equal("docx", OpenXmlIndexConstants.WordProcessingDocumentFileExtension);
    }

    [Fact]
    public void PresentationDocumentFileExtension_HasExpectedValue()
    {
        Assert.Equal("pptx", OpenXmlIndexConstants.PresentationDocumentFileExtension);
    }

    [Fact]
    public void SpreadsheetDocumentFileExtension_HasExpectedValue()
    {
        Assert.Equal("xlsx", OpenXmlIndexConstants.SpreadsheetDocumentFileExtension);
    }

    [Fact]
    public void OpenXmlCategory_HasExpectedValue()
    {
        Assert.Equal("openxml", OpenXmlIndexConstants.OpenXmlCategory);
    }

    [Fact]
    public void MaxExtractedContentLength_Is10MB()
    {
        Assert.Equal(10 * 1024 * 1024, OpenXmlIndexConstants.MaxExtractedContentLength);
    }

    [Fact]
    public void MaxFileSize_Is100MB()
    {
        Assert.Equal(100 * 1024 * 1024, OpenXmlIndexConstants.MaxFileSize);
    }

    [Fact]
    public void MaxSharedStringCount_Is1Million()
    {
        Assert.Equal(1_000_000, OpenXmlIndexConstants.MaxSharedStringCount);
    }

    [Fact]
    public void MaxCharactersInPart_Is10Million()
    {
        Assert.Equal(10_000_000, OpenXmlIndexConstants.MaxCharactersInPart);
    }
}
