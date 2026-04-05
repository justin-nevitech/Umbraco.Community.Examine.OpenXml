using Examine;
using Microsoft.Extensions.Logging;
using Moq;
using Umbraco.Cms.Core.IO;
using Umbraco.Cms.Core.Models;
using Umbraco.Cms.Core.Strings;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public class OpenXmlServiceTests
{
    private readonly Mock<IOpenXmlTextExtractorFactory> _factoryMock = new();
    private readonly Mock<ILogger<OpenXmlService>> _loggerMock = new();
    private readonly Mock<IOpenXmlTextExtractor> _extractorMock = new();

    private OpenXmlService CreateService(Mock<IFileSystem> fileSystemMock)
    {
        var mediaFileManager = new MediaFileManager(
            fileSystemMock.Object,
            Mock.Of<IMediaPathScheme>(),
            Mock.Of<ILogger<MediaFileManager>>(),
            Mock.Of<IShortStringHelper>(),
            Mock.Of<IServiceProvider>());
        return new OpenXmlService(
            _factoryMock.Object,
            mediaFileManager,
            _loggerMock.Object);
    }

    [Fact]
    public void ExtractOpenXml_WithValidDocxPath_ReturnsExtractedText()
    {
        var filePath = "/media/test.docx";
        var expectedText = "Hello World";
        var stream = new MemoryStream();

        var fileSystemMock = new Mock<IFileSystem>();
        fileSystemMock.Setup(f => f.OpenFile(filePath)).Returns(stream);
        _factoryMock.Setup(f => f.GetOpenXmlTextExtractor("docx")).Returns(_extractorMock.Object);
        _extractorMock.Setup(e => e.GetText(stream)).Returns(expectedText);

        var service = CreateService(fileSystemMock);
        var result = service.ExtractOpenXml(filePath);

        Assert.Equal(expectedText, result);
    }

    [Fact]
    public void ExtractOpenXml_WithEmptyExtension_ReturnsEmptyString()
    {
        var filePath = "/media/testfile";
        var fileSystemMock = new Mock<IFileSystem>();

        var service = CreateService(fileSystemMock);
        var result = service.ExtractOpenXml(filePath);

        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void ExtractOpenXml_WithNullStream_ReturnsEmptyString()
    {
        var filePath = "/media/test.docx";
        var fileSystemMock = new Mock<IFileSystem>();
        fileSystemMock.Setup(f => f.OpenFile(filePath)).Returns((Stream)null!);

        var service = CreateService(fileSystemMock);
        var result = service.ExtractOpenXml(filePath);

        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void ExtractOpenXml_CallsCorrectExtractorForPptx()
    {
        var filePath = "/media/presentation.pptx";
        var stream = new MemoryStream();

        var fileSystemMock = new Mock<IFileSystem>();
        fileSystemMock.Setup(f => f.OpenFile(filePath)).Returns(stream);
        _factoryMock.Setup(f => f.GetOpenXmlTextExtractor("pptx")).Returns(_extractorMock.Object);
        _extractorMock.Setup(e => e.GetText(stream)).Returns("slide text");

        var service = CreateService(fileSystemMock);
        service.ExtractOpenXml(filePath);

        _factoryMock.Verify(f => f.GetOpenXmlTextExtractor("pptx"), Times.Once);
    }

    [Fact]
    public void ExtractOpenXml_CallsCorrectExtractorForXlsx()
    {
        var filePath = "/media/spreadsheet.xlsx";
        var stream = new MemoryStream();

        var fileSystemMock = new Mock<IFileSystem>();
        fileSystemMock.Setup(f => f.OpenFile(filePath)).Returns(stream);
        _factoryMock.Setup(f => f.GetOpenXmlTextExtractor("xlsx")).Returns(_extractorMock.Object);
        _extractorMock.Setup(e => e.GetText(stream)).Returns("cell text");

        var service = CreateService(fileSystemMock);
        service.ExtractOpenXml(filePath);

        _factoryMock.Verify(f => f.GetOpenXmlTextExtractor("xlsx"), Times.Once);
    }

    [Fact]
    public void ExtractOpenXml_StripsLeadingDotFromExtension()
    {
        var filePath = "/media/file.docx";
        var stream = new MemoryStream();

        var fileSystemMock = new Mock<IFileSystem>();
        fileSystemMock.Setup(f => f.OpenFile(filePath)).Returns(stream);
        _factoryMock.Setup(f => f.GetOpenXmlTextExtractor("docx")).Returns(_extractorMock.Object);
        _extractorMock.Setup(e => e.GetText(stream)).Returns("content");

        var service = CreateService(fileSystemMock);
        service.ExtractOpenXml(filePath);

        _factoryMock.Verify(f => f.GetOpenXmlTextExtractor("docx"), Times.Once);
    }

    [Fact]
    public void ExtractOpenXml_FilePathWithJustDot_ReturnsEmptyString()
    {
        var filePath = "file.";
        var fileSystemMock = new Mock<IFileSystem>();

        var service = CreateService(fileSystemMock);
        var result = service.ExtractOpenXml(filePath);

        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void ExtractOpenXml_FileExceedsMaxSize_ReturnsEmptyString()
    {
        var filePath = "/media/huge.docx";
        var stream = new Mock<Stream>();
        stream.Setup(s => s.Length).Returns(OpenXmlIndexConstants.MaxFileSize + 1);
        stream.Setup(s => s.CanRead).Returns(true);

        var fileSystemMock = new Mock<IFileSystem>();
        fileSystemMock.Setup(f => f.OpenFile(filePath)).Returns(stream.Object);

        var service = CreateService(fileSystemMock);
        var result = service.ExtractOpenXml(filePath);

        Assert.Equal(string.Empty, result);
        _factoryMock.Verify(f => f.GetOpenXmlTextExtractor(It.IsAny<string>()), Times.Never);
    }

    [Fact]
    public void ExtractOpenXml_FileAtExactMaxSize_ProcessesNormally()
    {
        var filePath = "/media/exact.docx";
        var stream = new Mock<Stream>();
        stream.Setup(s => s.Length).Returns(OpenXmlIndexConstants.MaxFileSize); // exactly at limit, not over
        stream.Setup(s => s.CanRead).Returns(true);

        var fileSystemMock = new Mock<IFileSystem>();
        fileSystemMock.Setup(f => f.OpenFile(filePath)).Returns(stream.Object);
        _factoryMock.Setup(f => f.GetOpenXmlTextExtractor("docx")).Returns(_extractorMock.Object);
        _extractorMock.Setup(e => e.GetText(stream.Object)).Returns("content");

        var service = CreateService(fileSystemMock);
        var result = service.ExtractOpenXml(filePath);

        Assert.Equal("content", result);
    }

    [Fact]
    public void ExtractOpenXml_ExceptionDuringExtraction_ReturnsEmptyString()
    {
        var filePath = "/media/bad.docx";
        var stream = new MemoryStream();

        var fileSystemMock = new Mock<IFileSystem>();
        fileSystemMock.Setup(f => f.OpenFile(filePath)).Returns(stream);
        _factoryMock.Setup(f => f.GetOpenXmlTextExtractor("docx")).Returns(_extractorMock.Object);
        _extractorMock.Setup(e => e.GetText(stream)).Throws(new InvalidOperationException("Corrupted"));

        var service = CreateService(fileSystemMock);
        var result = service.ExtractOpenXml(filePath);

        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void ExtractOpenXml_ExceptionDuringExtraction_IsHandledByCallerInValueSetBuilder()
    {
        // Set up OpenXmlService to throw when extracting
        var filePath = "/media/bad.docx";
        var stream = new MemoryStream();

        var factoryMock = new Mock<IOpenXmlTextExtractorFactory>();
        var loggerMock = new Mock<ILogger<OpenXmlService>>();
        var extractorMock = new Mock<IOpenXmlTextExtractor>();
        extractorMock.Setup(e => e.GetText(It.IsAny<Stream>())).Throws(new InvalidOperationException("Corrupted file"));
        factoryMock.Setup(f => f.GetOpenXmlTextExtractor("docx")).Returns(extractorMock.Object);

        var fileSystemMock = new Mock<IFileSystem>();
        fileSystemMock.Setup(f => f.OpenFile(filePath)).Returns(stream);

        var mediaFileManager = new MediaFileManager(
            fileSystemMock.Object,
            Mock.Of<IMediaPathScheme>(),
            Mock.Of<ILogger<MediaFileManager>>(),
            Mock.Of<IShortStringHelper>(),
            Mock.Of<IServiceProvider>());

        var service = new OpenXmlService(factoryMock.Object, mediaFileManager, loggerMock.Object);

        var builderLoggerMock = new Mock<ILogger<OpenXmlIndexValueSetBuilder>>();
        var builder = new OpenXmlIndexValueSetBuilder(service, builderLoggerMock.Object);

        // Create a mock media item that returns the file path
        var mediaMock = new Mock<IMedia>();
        mediaMock.Setup(m => m.Id).Returns(1);
        mediaMock.Setup(m => m.Name).Returns("BadDoc");
        mediaMock.Setup(m => m.Path).Returns("-1,1");
        mediaMock.Setup(m => m.GetValue<string>(Umbraco.Cms.Core.Constants.Conventions.Media.File, null, null, false))
            .Returns(filePath);
        var contentTypeMock = new Mock<ISimpleContentType>();
        contentTypeMock.Setup(c => c.Alias).Returns("File");
        mediaMock.Setup(m => m.ContentType).Returns(contentTypeMock.Object);

        // The builder should catch the exception and continue (yielding a value set with empty content)
        var result = builder.GetValueSets(mediaMock.Object).ToList();

        // The builder catches the exception and continues, yielding a value set with empty extracted content
        Assert.Single(result);
        Assert.Equal(string.Empty, result[0].Values[OpenXmlIndexConstants.OpenXmlContentFieldName][0]);
    }
}
