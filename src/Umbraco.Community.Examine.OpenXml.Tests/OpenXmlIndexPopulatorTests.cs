using Examine;
using Moq;
using Umbraco.Cms.Core.Models;
using Umbraco.Cms.Core.Persistence.Querying;
using Umbraco.Cms.Core.Services;
using Umbraco.Cms.Infrastructure.Examine;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public class OpenXmlIndexPopulatorTests
{
    private readonly Mock<IMediaService> _mediaServiceMock = new();
    private readonly Mock<IOpenXmlIndexValueSetBuilder> _valueSetBuilderMock = new();
    private readonly Mock<IExamineManager> _examineManagerMock = new();
    private readonly Mock<IIndex> _indexMock = new();

    private OpenXmlIndexPopulator CreatePopulator(int? parentId = null)
    {
        IIndex? index = _indexMock.Object;
        _examineManagerMock.Setup(e => e.TryGetIndex(OpenXmlIndexConstants.OpenXmlIndexName, out index)).Returns(true);
        _indexMock.Setup(i => i.Name).Returns(OpenXmlIndexConstants.OpenXmlIndexName);

        return parentId.HasValue
            ? new OpenXmlIndexPopulator(parentId, _mediaServiceMock.Object, _valueSetBuilderMock.Object, _examineManagerMock.Object)
            : new OpenXmlIndexPopulator(_mediaServiceMock.Object, _valueSetBuilderMock.Object, _examineManagerMock.Object);
    }

    private static Mock<IMedia> CreateMediaMock(int id, string extension)
    {
        var mock = new Mock<IMedia>();
        mock.Setup(m => m.Id).Returns(id);
        mock.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns(extension);
        return mock;
    }

    // --- AddToIndex tests ---

    [Fact]
    public void AddToIndex_WithDocxMedia_IndexesItem()
    {
        var populator = CreatePopulator();
        var media = CreateMediaMock(1, "docx");

        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Returns(new[] { new ValueSet("1", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        populator.AddToIndex(media.Object);

        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.Once);
    }

    [Fact]
    public void AddToIndex_WithNonOpenXmlMedia_DoesNotIndex()
    {
        var populator = CreatePopulator();
        var media = CreateMediaMock(1, "jpg");

        populator.AddToIndex(media.Object);

        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.Never);
    }

    [Fact]
    public void AddToIndex_WhenIndexNotFound_DoesNotThrow()
    {
        IIndex? index = null;
        _examineManagerMock.Setup(e => e.TryGetIndex(OpenXmlIndexConstants.OpenXmlIndexName, out index)).Returns(false);
        var populator = new OpenXmlIndexPopulator(_mediaServiceMock.Object, _valueSetBuilderMock.Object, _examineManagerMock.Object);
        var media = CreateMediaMock(1, "docx");

        populator.AddToIndex(media.Object);

        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.Never);
    }

    // --- RemoveFromIndex tests ---

    [Fact]
    public void RemoveFromIndex_ByMedia_DeletesFromIndex()
    {
        var populator = CreatePopulator();
        var media = CreateMediaMock(42, "docx");

        populator.RemoveFromIndex(media.Object);

        _indexMock.Verify(i => i.DeleteFromIndex(It.Is<IEnumerable<string>>(ids => ids.Contains("42"))), Times.Once);
    }

    [Fact]
    public void RemoveFromIndex_ByIds_DeletesFromIndex()
    {
        var populator = CreatePopulator();

        populator.RemoveFromIndex(42, 43);

        _indexMock.Verify(i => i.DeleteFromIndex(It.Is<IEnumerable<string>>(ids => ids.Contains("42") && ids.Contains("43"))), Times.Once);
    }

    [Fact]
    public void RemoveFromIndex_WhenIndexNotFound_DoesNotThrow()
    {
        IIndex? index = null;
        _examineManagerMock.Setup(e => e.TryGetIndex(OpenXmlIndexConstants.OpenXmlIndexName, out index)).Returns(false);
        var populator = new OpenXmlIndexPopulator(_mediaServiceMock.Object, _valueSetBuilderMock.Object, _examineManagerMock.Object);

        populator.RemoveFromIndex(42);

        _indexMock.Verify(i => i.DeleteFromIndex(It.IsAny<IEnumerable<string>>()), Times.Never);
    }

    [Fact]
    public void AddToIndex_WithAllSupportedExtensions_IndexesAll()
    {
        var populator = CreatePopulator();
        var docx = CreateMediaMock(1, "docx");
        var pptx = CreateMediaMock(2, "pptx");
        var xlsx = CreateMediaMock(3, "xlsx");

        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Returns(new[] { new ValueSet("1", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        populator.AddToIndex(docx.Object, pptx.Object, xlsx.Object);

        _valueSetBuilderMock.Verify(b => b.GetValueSets(It.Is<IMedia[]>(arr => arr.Length == 3)), Times.Once);
    }

    [Fact]
    public void AddToIndex_WithNullExtension_DoesNotThrow()
    {
        var populator = CreatePopulator();
        var media = CreateMediaMock(1, null!);

        populator.AddToIndex(media.Object);

        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.Never);
    }

    [Fact]
    public void AddToIndex_WithMixOfSupportedAndUnsupported_OnlyIndexesSupported()
    {
        var populator = CreatePopulator();
        var docx = CreateMediaMock(1, "docx");
        var jpg = CreateMediaMock(2, "jpg");
        var xlsx = CreateMediaMock(3, "xlsx");

        IMedia[]? capturedMedia = null;
        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Callback<IMedia[]>(media => capturedMedia = media)
            .Returns(new[] { new ValueSet("1", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        populator.AddToIndex(docx.Object, jpg.Object, xlsx.Object);

        Assert.NotNull(capturedMedia);
        Assert.Equal(2, capturedMedia.Length);
        Assert.Contains(capturedMedia, m => m.Id == 1);
        Assert.Contains(capturedMedia, m => m.Id == 3);
    }

    // --- PopulateIndexes tests (via Populate) ---

    [Fact]
    public void PopulateIndexes_WithMultiplePages_IndexesAllPages()
    {
        var populator = CreatePopulator();

        var docxMedia1 = CreateMediaMock(1, "docx");
        var jpgMedia = CreateMediaMock(2, "jpg");
        var docxMedia2 = CreateMediaMock(3, "docx");

        // Page 0: total = 10001 forces a second page (pageSize is 10000)
        long totalRecordsPage0 = 10001;
        _mediaServiceMock.Setup(m => m.GetPagedDescendants(-1, 0, 10000, out totalRecordsPage0,
                It.IsAny<IQuery<IMedia>?>(), It.IsAny<Ordering?>()))
            .Returns(new[] { docxMedia1.Object, jpgMedia.Object });

        // Page 1: same total
        long totalRecordsPage1 = 10001;
        _mediaServiceMock.Setup(m => m.GetPagedDescendants(-1, 1, 10000, out totalRecordsPage1,
                It.IsAny<IQuery<IMedia>?>(), It.IsAny<Ordering?>()))
            .Returns(new[] { docxMedia2.Object });

        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Returns(new[] { new ValueSet("1", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        populator.Populate(_indexMock.Object);

        // Should have fetched both pages
        long verifyTotal;
        _mediaServiceMock.Verify(m => m.GetPagedDescendants(-1, 0, 10000, out verifyTotal,
            It.IsAny<IQuery<IMedia>?>(), It.IsAny<Ordering?>()), Times.Once);
        _mediaServiceMock.Verify(m => m.GetPagedDescendants(-1, 1, 10000, out verifyTotal,
            It.IsAny<IQuery<IMedia>?>(), It.IsAny<Ordering?>()), Times.Once);
    }

    [Fact]
    public void PopulateIndexes_FiltersNonOpenXmlMedia()
    {
        var populator = CreatePopulator();

        var docxMedia = CreateMediaMock(1, "docx");
        var jpgMedia = CreateMediaMock(2, "jpg");
        var pdfMedia = CreateMediaMock(3, "pdf");

        long totalRecords = 3; // less than pageSize, so single page
        _mediaServiceMock.Setup(m => m.GetPagedDescendants(-1, 0, 10000, out totalRecords,
                It.IsAny<IQuery<IMedia>?>(), It.IsAny<Ordering?>()))
            .Returns(new[] { docxMedia.Object, jpgMedia.Object, pdfMedia.Object });

        IMedia[]? capturedMedia = null;
        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Callback<IMedia[]>(media => capturedMedia = media)
            .Returns(new[] { new ValueSet("1", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        populator.Populate(_indexMock.Object);

        // Only the docx media should be passed to the value set builder
        Assert.NotNull(capturedMedia);
        Assert.Single(capturedMedia);
        Assert.Equal(1, capturedMedia[0].Id);
    }

    [Fact]
    public void PopulateIndexes_WithCustomParentId_UsesCorrectParentId()
    {
        var populator = CreatePopulator(parentId: 100);

        long totalRecords = 0;
        _mediaServiceMock.Setup(m => m.GetPagedDescendants(100, 0, 10000, out totalRecords,
                It.IsAny<IQuery<IMedia>?>(), It.IsAny<Ordering?>()))
            .Returns(Array.Empty<IMedia>());

        populator.Populate(_indexMock.Object);

        // Verify parentId 100 was used (not -1)
        _mediaServiceMock.Verify(m => m.GetPagedDescendants(100, 0, 10000, out totalRecords,
            It.IsAny<IQuery<IMedia>?>(), It.IsAny<Ordering?>()), Times.Once);
    }

    [Fact]
    public void PopulateIndexes_WithNoMedia_DoesNotIndex()
    {
        var populator = CreatePopulator();

        long totalRecords = 0;
        _mediaServiceMock.Setup(m => m.GetPagedDescendants(-1, 0, 10000, out totalRecords,
                It.IsAny<IQuery<IMedia>?>(), It.IsAny<Ordering?>()))
            .Returns(Array.Empty<IMedia>());

        populator.Populate(_indexMock.Object);

        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.Never);
    }
}
