using Microsoft.Extensions.Logging;
using Moq;
using Umbraco.Cms.Core;
using Umbraco.Cms.Core.IO;
using Umbraco.Cms.Core.Models;
using Umbraco.Cms.Core.Strings;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public class OpenXmlIndexValueSetBuilderTests
{
    private readonly Mock<ILogger<OpenXmlIndexValueSetBuilder>> _loggerMock = new();
    private readonly OpenXmlIndexValueSetBuilder _builder;

    public OpenXmlIndexValueSetBuilderTests()
    {
        var factoryMock = new Mock<IOpenXmlTextExtractorFactory>();
        var fileSystemMock = new Mock<IFileSystem>();

        var mediaFileManager = new MediaFileManager(
            fileSystemMock.Object,
            Mock.Of<IMediaPathScheme>(),
            Mock.Of<ILogger<MediaFileManager>>(),
            Mock.Of<IShortStringHelper>(),
            Mock.Of<IServiceProvider>());

        var serviceLoggerMock = new Mock<ILogger<OpenXmlService>>();

        var extractorMock = new Mock<IOpenXmlTextExtractor>();
        extractorMock.Setup(e => e.GetText(It.IsAny<Stream>())).Returns("extracted content");
        factoryMock.Setup(f => f.GetOpenXmlTextExtractor(It.IsAny<string>())).Returns(extractorMock.Object);

        var stream = new MemoryStream();
        fileSystemMock.Setup(f => f.OpenFile(It.IsAny<string>())).Returns(stream);

        var service = new OpenXmlService(factoryMock.Object, mediaFileManager, serviceLoggerMock.Object);
        _builder = new OpenXmlIndexValueSetBuilder(service, _loggerMock.Object);
    }

    [Fact]
    public void GetValueSets_WithValidMedia_ReturnsValueSet()
    {
        var media = CreateMockMedia(1, "TestDoc", "/media/test.docx", "-1,1");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Single(result);
    }

    [Fact]
    public void GetValueSets_ValueSetContainsNodeName()
    {
        var media = CreateMockMedia(42, "My Document", "/media/mydoc.docx", "-1,42");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Single(result);
        var valueSet = result[0];
        Assert.True(valueSet.Values.ContainsKey("nodeName"));
        Assert.Equal("My Document", valueSet.Values["nodeName"][0]);
    }

    [Fact]
    public void GetValueSets_ValueSetContainsId()
    {
        var media = CreateMockMedia(42, "My Document", "/media/mydoc.docx", "-1,42");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Single(result);
        var valueSet = result[0];
        Assert.True(valueSet.Values.ContainsKey("id"));
        Assert.Equal(42, valueSet.Values["id"][0]);
    }

    [Fact]
    public void GetValueSets_ValueSetContainsPath()
    {
        var media = CreateMockMedia(42, "My Document", "/media/mydoc.docx", "-1,42");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Single(result);
        var valueSet = result[0];
        Assert.True(valueSet.Values.ContainsKey("path"));
        Assert.Equal("-1,42", valueSet.Values["path"][0]);
    }

    [Fact]
    public void GetValueSets_ValueSetContainsFileTextContent()
    {
        var media = CreateMockMedia(42, "My Document", "/media/mydoc.docx", "-1,42");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Single(result);
        var valueSet = result[0];
        Assert.True(valueSet.Values.ContainsKey(OpenXmlIndexConstants.OpenXmlContentFieldName));
    }

    [Fact]
    public void GetValueSets_ValueSetHasCorrectCategory()
    {
        var media = CreateMockMedia(42, "My Document", "/media/mydoc.docx", "-1,42");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Single(result);
        Assert.Equal(OpenXmlIndexConstants.OpenXmlCategory, result[0].Category);
    }

    [Fact]
    public void GetValueSets_ValueSetIdIsMediaId()
    {
        var media = CreateMockMedia(42, "My Document", "/media/mydoc.docx", "-1,42");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Single(result);
        Assert.Equal("42", result[0].Id);
    }

    [Fact]
    public void GetValueSets_MediaWithoutUmbracoFile_IsSkipped()
    {
        var media = CreateMockMedia(1, "NoFile", null, "-1,1");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Empty(result);
    }

    [Fact]
    public void GetValueSets_MediaWithEmptyUmbracoFile_IsSkipped()
    {
        var media = CreateMockMedia(1, "EmptyFile", "", "-1,1");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Empty(result);
    }

    [Fact]
    public void GetValueSets_MediaWithWhitespaceUmbracoFile_IsSkipped()
    {
        var media = CreateMockMedia(1, "WhitespaceFile", "   ", "-1,1");

        var result = _builder.GetValueSets(media).ToList();

        Assert.Empty(result);
    }

    [Fact]
    public void GetValueSets_MultipleMedia_ReturnsMultipleValueSets()
    {
        var media1 = CreateMockMedia(1, "Doc1", "/media/doc1.docx", "-1,1");
        var media2 = CreateMockMedia(2, "Doc2", "/media/doc2.docx", "-1,2");

        var result = _builder.GetValueSets(media1, media2).ToList();

        Assert.Equal(2, result.Count);
    }

    private static IMedia CreateMockMedia(int id, string name, string? umbracoFilePath, string path)
    {
        var mediaMock = new Mock<IMedia>();
        mediaMock.Setup(m => m.Id).Returns(id);
        mediaMock.Setup(m => m.Name).Returns(name);
        mediaMock.Setup(m => m.Path).Returns(path);
        mediaMock.Setup(m => m.GetValue<string>(Constants.Conventions.Media.File, null, null, false))
            .Returns(umbracoFilePath!);

        var contentTypeMock = new Mock<ISimpleContentType>();
        contentTypeMock.Setup(c => c.Alias).Returns("File");
        mediaMock.Setup(m => m.ContentType).Returns(contentTypeMock.Object);

        return mediaMock.Object;
    }
}
