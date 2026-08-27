using Microsoft.Extensions.Logging;
using Moq;
using Umbraco.Cms.Core.DependencyInjection;
using Umbraco.Cms.Core.IO;
using Umbraco.Cms.Core.Scoping;
using Umbraco.Cms.Core.Strings;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public static class TestHelper
{
    public static string GetTestFilePath(string fileName)
    {
        return Path.Combine(AppContext.BaseDirectory, "TestFiles", fileName);
    }

    public static Stream GetTestFileStream(string fileName)
    {
        var path = GetTestFilePath(fileName);
        return File.OpenRead(path);
    }

    /// <summary>
    ///     Builds a <see cref="MediaFileManager" /> over the supplied file system.
    /// </summary>
    /// <remarks>
    ///     From Umbraco 18 the public constructor resolves <c>Lazy&lt;ICoreScopeProvider&gt;</c> from
    ///     <see cref="StaticServiceProvider" />, which Umbraco populates during boot. These tests never
    ///     boot Umbraco, so the locator is seeded here with a stub. On Umbraco 13 and 17 the constructor
    ///     does not touch the locator and the assignment is simply inert, which keeps this one helper
    ///     valid for every supported major.
    /// </remarks>
    public static MediaFileManager CreateMediaFileManager(IFileSystem fileSystem)
    {
        StaticServiceProvider.Instance = StubServiceProvider.Instance;

        return new MediaFileManager(
            fileSystem,
            Mock.Of<IMediaPathScheme>(),
            Mock.Of<ILogger<MediaFileManager>>(),
            Mock.Of<IShortStringHelper>(),
            StubServiceProvider.Instance);
    }

    /// <summary>
    ///     Resolves only what <see cref="MediaFileManager" /> asks of the service locator.
    /// </summary>
    private sealed class StubServiceProvider : IServiceProvider
    {
        public static readonly StubServiceProvider Instance = new();

        private readonly Lazy<ICoreScopeProvider> _coreScopeProvider = new(Mock.Of<ICoreScopeProvider>);

        public object? GetService(Type serviceType) =>
            serviceType == typeof(Lazy<ICoreScopeProvider>) ? _coreScopeProvider : null;
    }
}
