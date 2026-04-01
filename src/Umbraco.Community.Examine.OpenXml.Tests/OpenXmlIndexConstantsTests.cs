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
}
