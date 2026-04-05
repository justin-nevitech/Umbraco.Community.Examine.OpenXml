using Examine;
using Microsoft.Extensions.Logging;
using Moq;
using Umbraco.Cms.Core;
using Umbraco.Cms.Core.Cache;
using Umbraco.Cms.Core.Models;
using Umbraco.Cms.Core.Persistence.Querying;
using Umbraco.Cms.Core.Services;
using Umbraco.Cms.Core.Services.Changes;
using Umbraco.Cms.Core.Sync;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public class OpenXmlCacheNotificationHandlerTests
{
    private readonly Mock<IMediaService> _mediaServiceMock = new();
    private readonly Mock<IRuntimeState> _runtimeStateMock = new();
    private readonly Mock<ILogger<OpenXmlCacheNotificationHandler>> _loggerMock = new();
    private readonly Mock<IExamineManager> _examineManagerMock = new();
    private readonly Mock<IOpenXmlIndexValueSetBuilder> _valueSetBuilderMock = new();
    private readonly Mock<IIndex> _indexMock = new();

    private OpenXmlCacheNotificationHandler CreateHandler()
    {
        IIndex? index = _indexMock.Object;
        _examineManagerMock.Setup(e => e.TryGetIndex(OpenXmlIndexConstants.OpenXmlIndexName, out index)).Returns(true);

        var populator = new OpenXmlIndexPopulator(
            _mediaServiceMock.Object,
            _valueSetBuilderMock.Object,
            _examineManagerMock.Object);

        return new OpenXmlCacheNotificationHandler(
            _mediaServiceMock.Object,
            populator,
            _runtimeStateMock.Object,
            _loggerMock.Object);
    }

    private static MediaCacheRefresher.JsonPayload CreatePayload(int id, TreeChangeTypes changeTypes)
    {
        return new MediaCacheRefresher.JsonPayload(id, null, changeTypes);
    }

    [Fact]
    public void Handle_NonRunRuntimeLevel_SkipsProcessing()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Boot);
        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(1, TreeChangeTypes.RefreshNode) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        _mediaServiceMock.Verify(m => m.GetById(It.IsAny<int>()), Times.Never);
    }

    [Fact]
    public void Handle_UnsupportedMessageType_LogsWarningAndReturns()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);
        var handler = CreateHandler();
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(new object(), MessageType.RefreshAll);

        handler.Handle(notification);

        _mediaServiceMock.Verify(m => m.GetById(It.IsAny<int>()), Times.Never);
    }

    [Fact]
    public void Handle_RemoveChangeType_RemovesFromIndex()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);
        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(42, TreeChangeTypes.Remove) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        _indexMock.Verify(i => i.DeleteFromIndex(It.Is<IEnumerable<string>>(ids => ids.Contains("42"))), Times.Once);
    }

    [Fact]
    public void Handle_TrashedMedia_RemovesFromIndex()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);
        var mediaMock = new Mock<IMedia>();
        mediaMock.Setup(m => m.Id).Returns(42);
        mediaMock.Setup(m => m.Trashed).Returns(true);
        _mediaServiceMock.Setup(m => m.GetById(42)).Returns(mediaMock.Object);

        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(42, TreeChangeTypes.RefreshNode) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        _indexMock.Verify(i => i.DeleteFromIndex(It.Is<IEnumerable<string>>(ids => ids.Contains("42"))), Times.Once);
    }

    [Fact]
    public void Handle_NonTrashedMedia_AddsToIndex()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);
        var mediaMock = new Mock<IMedia>();
        mediaMock.Setup(m => m.Id).Returns(42);
        mediaMock.Setup(m => m.Trashed).Returns(false);
        mediaMock.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");
        _mediaServiceMock.Setup(m => m.GetById(42)).Returns(mediaMock.Object);
        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Returns(new[] { new ValueSet("42", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(42, TreeChangeTypes.RefreshNode) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.Once);
    }

    [Fact]
    public void Handle_MediaNotFound_RemovesFromIndex()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);
        _mediaServiceMock.Setup(m => m.GetById(42)).Returns((IMedia?)null);

        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(42, TreeChangeTypes.RefreshNode) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        _indexMock.Verify(i => i.DeleteFromIndex(It.Is<IEnumerable<string>>(ids => ids.Contains("42"))), Times.Once);
    }

    [Fact]
    public void Handle_RefreshAll_SkipsWithoutError()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);
        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(0, TreeChangeTypes.RefreshAll) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        // Should not throw
        handler.Handle(notification);

        _mediaServiceMock.Verify(m => m.GetById(It.IsAny<int>()), Times.Never);
        _indexMock.Verify(i => i.DeleteFromIndex(It.IsAny<IEnumerable<string>>()), Times.Never);
        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.Never);
    }

    [Fact]
    public void Handle_RefreshBranchWithDescendants_ProcessesAllDescendants()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);

        var parentMedia = new Mock<IMedia>();
        parentMedia.Setup(m => m.Id).Returns(10);
        parentMedia.Setup(m => m.Trashed).Returns(false);
        parentMedia.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");
        _mediaServiceMock.Setup(m => m.GetById(10)).Returns(parentMedia.Object);

        var descendant1 = new Mock<IMedia>();
        descendant1.Setup(m => m.Id).Returns(11);
        descendant1.Setup(m => m.Trashed).Returns(false);
        descendant1.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");

        var descendant2 = new Mock<IMedia>();
        descendant2.Setup(m => m.Id).Returns(12);
        descendant2.Setup(m => m.Trashed).Returns(false);
        descendant2.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");

        long totalRecords = 2;
        _mediaServiceMock.Setup(m => m.GetPagedDescendants(10, 0, 500, out totalRecords, It.IsAny<Umbraco.Cms.Core.Persistence.Querying.IQuery<IMedia>?>(), It.IsAny<Umbraco.Cms.Core.Services.Ordering?>()))
            .Returns(new[] { descendant1.Object, descendant2.Object });

        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Returns(new[] { new ValueSet("10", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(10, TreeChangeTypes.RefreshBranch) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        // Parent should be added to index
        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.AtLeastOnce);
        // Descendants should be processed
        _mediaServiceMock.Verify(m => m.GetPagedDescendants(10, 0, 500, out totalRecords, It.IsAny<Umbraco.Cms.Core.Persistence.Querying.IQuery<IMedia>?>(), It.IsAny<Umbraco.Cms.Core.Services.Ordering?>()), Times.Once);
    }

    [Fact]
    public void Handle_RefreshBranchWithTrashedDescendants_RemovesTrashedFromIndex()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);

        var parentMedia = new Mock<IMedia>();
        parentMedia.Setup(m => m.Id).Returns(10);
        parentMedia.Setup(m => m.Trashed).Returns(false);
        parentMedia.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");
        _mediaServiceMock.Setup(m => m.GetById(10)).Returns(parentMedia.Object);

        var trashedDescendant = new Mock<IMedia>();
        trashedDescendant.Setup(m => m.Id).Returns(20);
        trashedDescendant.Setup(m => m.Trashed).Returns(true);

        var nonTrashedDescendant = new Mock<IMedia>();
        nonTrashedDescendant.Setup(m => m.Id).Returns(21);
        nonTrashedDescendant.Setup(m => m.Trashed).Returns(false);
        nonTrashedDescendant.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");

        long totalRecords = 2;
        _mediaServiceMock.Setup(m => m.GetPagedDescendants(10, 0, 500, out totalRecords, It.IsAny<Umbraco.Cms.Core.Persistence.Querying.IQuery<IMedia>?>(), It.IsAny<Umbraco.Cms.Core.Services.Ordering?>()))
            .Returns(new[] { trashedDescendant.Object, nonTrashedDescendant.Object });

        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Returns(new[] { new ValueSet("10", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(10, TreeChangeTypes.RefreshBranch) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        // Trashed descendant should be removed
        _indexMock.Verify(i => i.DeleteFromIndex(It.Is<IEnumerable<string>>(ids => ids.Contains("20"))), Times.Once);
        // Non-trashed descendant should be added
        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.AtLeastOnce);
    }

    [Fact]
    public void Handle_MultiplePayloadsInNotification_ProcessesAll()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);

        var media1 = new Mock<IMedia>();
        media1.Setup(m => m.Id).Returns(1);
        media1.Setup(m => m.Trashed).Returns(false);
        media1.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");
        _mediaServiceMock.Setup(m => m.GetById(1)).Returns(media1.Object);

        var media2 = new Mock<IMedia>();
        media2.Setup(m => m.Id).Returns(2);
        media2.Setup(m => m.Trashed).Returns(true);
        _mediaServiceMock.Setup(m => m.GetById(2)).Returns(media2.Object);

        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Returns(new[] { new ValueSet("1", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        var handler = CreateHandler();
        var payloads = new[]
        {
            CreatePayload(1, TreeChangeTypes.RefreshNode),
            CreatePayload(2, TreeChangeTypes.RefreshNode)
        };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        // Both media items should be looked up
        _mediaServiceMock.Verify(m => m.GetById(1), Times.Once);
        _mediaServiceMock.Verify(m => m.GetById(2), Times.Once);
        // Media 1 (not trashed) should be indexed
        _indexMock.Verify(i => i.IndexItems(It.IsAny<IEnumerable<ValueSet>>()), Times.Once);
        // Media 2 (trashed) should be removed
        _indexMock.Verify(i => i.DeleteFromIndex(It.Is<IEnumerable<string>>(ids => ids.Contains("2"))), Times.Once);
    }

    [Fact]
    public void Handle_RefreshByPayloadWithWrongObjectType_LogsWarningAndReturns()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);
        var handler = CreateHandler();
        // RefreshByPayload message type but with a non-JsonPayload[] object
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification("not a payload array", MessageType.RefreshByPayload);

        handler.Handle(notification);

        _mediaServiceMock.Verify(m => m.GetById(It.IsAny<int>()), Times.Never);
    }

    [Fact]
    public void Handle_ExceptionDuringProcessing_DoesNotThrow()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);
        _mediaServiceMock.Setup(m => m.GetById(It.IsAny<int>())).Throws(new InvalidOperationException("Database error"));

        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(42, TreeChangeTypes.RefreshNode) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        // Should not throw — handler catches and logs
        var exception = Record.Exception(() => handler.Handle(notification));

        Assert.Null(exception);
    }

    [Fact]
    public void Handle_RefreshBranchWithMultiplePages_ProcessesAllPages()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);

        var parentMedia = new Mock<IMedia>();
        parentMedia.Setup(m => m.Id).Returns(10);
        parentMedia.Setup(m => m.Trashed).Returns(false);
        parentMedia.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");
        _mediaServiceMock.Setup(m => m.GetById(10)).Returns(parentMedia.Object);

        var descendant1 = new Mock<IMedia>();
        descendant1.Setup(m => m.Id).Returns(11);
        descendant1.Setup(m => m.Trashed).Returns(false);
        descendant1.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");

        var descendant2 = new Mock<IMedia>();
        descendant2.Setup(m => m.Id).Returns(12);
        descendant2.Setup(m => m.Trashed).Returns(false);
        descendant2.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");

        // Page 0: 1 descendant, total = 501 (more than one page of 500)
        long totalRecords = 501;
        _mediaServiceMock.Setup(m => m.GetPagedDescendants(10, 0, 500, out totalRecords,
                It.IsAny<Umbraco.Cms.Core.Persistence.Querying.IQuery<IMedia>?>(),
                It.IsAny<Umbraco.Cms.Core.Services.Ordering?>()))
            .Returns(new[] { descendant1.Object });

        // Page 1: 1 descendant
        _mediaServiceMock.Setup(m => m.GetPagedDescendants(10, 1, 500, out totalRecords,
                It.IsAny<Umbraco.Cms.Core.Persistence.Querying.IQuery<IMedia>?>(),
                It.IsAny<Umbraco.Cms.Core.Services.Ordering?>()))
            .Returns(new[] { descendant2.Object });

        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Returns(new[] { new ValueSet("10", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        var handler = CreateHandler();
        var payloads = new[] { CreatePayload(10, TreeChangeTypes.RefreshBranch) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        // Should have fetched both pages
        _mediaServiceMock.Verify(m => m.GetPagedDescendants(10, 0, 500, out totalRecords,
            It.IsAny<Umbraco.Cms.Core.Persistence.Querying.IQuery<IMedia>?>(),
            It.IsAny<Umbraco.Cms.Core.Services.Ordering?>()), Times.Once);
        _mediaServiceMock.Verify(m => m.GetPagedDescendants(10, 1, 500, out totalRecords,
            It.IsAny<Umbraco.Cms.Core.Persistence.Querying.IQuery<IMedia>?>(),
            It.IsAny<Umbraco.Cms.Core.Services.Ordering?>()), Times.Once);
    }

    [Fact]
    public void Handle_RefreshNodeWithoutBranch_DoesNotProcessDescendants()
    {
        _runtimeStateMock.Setup(r => r.Level).Returns(RuntimeLevel.Run);

        var media = new Mock<IMedia>();
        media.Setup(m => m.Id).Returns(42);
        media.Setup(m => m.Trashed).Returns(false);
        media.Setup(m => m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias, null, null, false))
            .Returns("docx");
        _mediaServiceMock.Setup(m => m.GetById(42)).Returns(media.Object);

        _valueSetBuilderMock.Setup(b => b.GetValueSets(It.IsAny<IMedia[]>()))
            .Returns(new[] { new ValueSet("42", "openxml", "File", new Dictionary<string, object> { ["test"] = "value" }) });

        var handler = CreateHandler();
        // RefreshNode only, NOT RefreshBranch
        var payloads = new[] { CreatePayload(42, TreeChangeTypes.RefreshNode) };
        var notification = new Umbraco.Cms.Core.Notifications.MediaCacheRefresherNotification(payloads, MessageType.RefreshByPayload);

        handler.Handle(notification);

        long totalRecords;
        _mediaServiceMock.Verify(m => m.GetPagedDescendants(It.IsAny<int>(), It.IsAny<int>(), It.IsAny<int>(), out totalRecords, It.IsAny<Umbraco.Cms.Core.Persistence.Querying.IQuery<IMedia>?>(), It.IsAny<Umbraco.Cms.Core.Services.Ordering?>()), Times.Never);
    }
}
