using Examine;
using Microsoft.Extensions.DependencyInjection;
using Umbraco.Cms.Core.DependencyInjection;
using Umbraco.Cms.Core.Notifications;
using Umbraco.Cms.Infrastructure.Examine;
using Umbraco.Extensions;

namespace Umbraco.Community.Examine.OpenXml
{
    public static class BuilderExtensions
    {
        public static IUmbracoBuilder AddExamineOpenXml(this IUmbracoBuilder builder)
        {
            if (builder.Services.Any(x => x.ServiceType == typeof(IOpenXmlTextExtractorFactory)))
            {
                // Assume that Examine.OpenXml is already composed if any implementation of IOpenXmlTextExtractorFactory is registered.
                return builder;
            }

            //Register the services used to make this all work
            builder.Services.AddUnique<IWordProcessingDocumentTextExtractor, WordProcessingDocumentTextExtractor>();
            builder.Services.AddUnique<IPresentationDocumentTextExtractor, PresentationDocumentTextExtractor>();
            builder.Services.AddUnique<ISpreadsheetDocumentTextExtractor, SpreadsheetDocumentTextExtractor>();
            builder.Services.AddUnique<IOpenXmlTextExtractorFactory, OpenXmlTextExtractorFactory>();
            builder.Services.AddSingleton<OpenXmlService>();
            builder.Services.AddUnique<IOpenXmlIndexValueSetBuilder, OpenXmlIndexValueSetBuilder>();
            builder.Services.AddSingleton<IIndexPopulator, OpenXmlIndexPopulator>();
            builder.Services.AddSingleton<OpenXmlIndexPopulator>();

            builder.Services
                .AddExamineLuceneIndex<OpenXmlLuceneIndex, ConfigurationEnabledDirectoryFactory>(OpenXmlIndexConstants.OpenXmlIndexName)
                .ConfigureOptions<ConfigureOpenXmlIndexOptions>();

            builder.AddNotificationHandler<MediaCacheRefresherNotification, OpenXmlCacheNotificationHandler>();

            return builder;
        }
    }
}
