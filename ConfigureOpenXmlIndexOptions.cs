using Examine;
using Examine.Lucene;
using Lucene.Net.Analysis.Standard;
using Lucene.Net.Index;
using Lucene.Net.Util;
using Microsoft.Extensions.Options;
using Umbraco.Cms.Core.Configuration.Models;

namespace UmbracoExamine.OpenXml
{
    // See ConfigureIndexOptions in the CMS to see how this is used.
    /// <summary>
    /// Configures the index options to construct the Examine OpenXml Index
    /// </summary>
    public class ConfigureOpenXmlIndexOptions : IConfigureNamedOptions<LuceneDirectoryIndexOptions>
    {
        private readonly IOptions<IndexCreatorSettings> _settings;

        public ConfigureOpenXmlIndexOptions(IOptions<IndexCreatorSettings> settings)
        {
            _settings = settings;
        }

        public void Configure(string? name, LuceneDirectoryIndexOptions options)
        {
            if (name?.Equals(OpenXmlIndexConstants.OpenXmlIndexName) ?? false)
            {
                options.Analyzer = new StandardAnalyzer(LuceneVersion.LUCENE_48);
                options.Validator = new OpenXmlValueSetValidator(true, null);
                options.FieldDefinitions = new FieldDefinitionCollection(
                    new FieldDefinition(OpenXmlIndexConstants.OpenXmlContentFieldName, FieldDefinitionTypes.FullText));

                // Ensures indexes are unlocked on startup
                options.UnlockIndex = true;

                if (_settings.Value.LuceneDirectoryFactory == LuceneDirectoryFactory.SyncedTempFileSystemDirectoryFactory)
                {
                    // if this directory factory is enabled then a snapshot deletion policy is required
                    options.IndexDeletionPolicy = new SnapshotDeletionPolicy(new KeepOnlyLastCommitDeletionPolicy());
                }
            }
        }

        public void Configure(LuceneDirectoryIndexOptions options)
        {
            throw new System.NotImplementedException("This is never called and is just part of the interface");
        }
    }
}