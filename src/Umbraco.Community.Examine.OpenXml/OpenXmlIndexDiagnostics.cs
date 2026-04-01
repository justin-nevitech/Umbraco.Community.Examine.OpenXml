using Examine.Lucene;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Umbraco.Cms.Core.Hosting;
using Umbraco.Cms.Infrastructure.Examine;

namespace Umbraco.Community.Examine.OpenXml
{
    internal class OpenXmlIndexDiagnostics : LuceneIndexDiagnostics
    {
        private readonly OpenXmlLuceneIndex _index;
        public OpenXmlIndexDiagnostics(
            OpenXmlLuceneIndex index,
            ILogger<LuceneIndexDiagnostics> logger,
            IHostingEnvironment hostingEnvironment,
            IOptionsMonitor<LuceneDirectoryIndexOptions> indexOptions)
            : base(index, logger, hostingEnvironment, indexOptions)
        {
            _index = index;
        }


        public override IReadOnlyDictionary<string, object?> Metadata
        {
            get
            {
                var d = base.Metadata.ToDictionary(x => x.Key, x => x.Value);

                if (_index.ValueSetValidator is ValueSetValidator vsv)
                {
                    d[nameof(ValueSetValidator.IncludeItemTypes)] = vsv.IncludeItemTypes;
                    d[nameof(ContentValueSetValidator.ExcludeItemTypes)] = vsv.ExcludeItemTypes;
                }

                if (_index.ValueSetValidator is OpenXmlValueSetValidator cvsv)
                {
                    d[nameof(ContentValueSetValidator.ParentId)] = cvsv.ParentId;
                }

                return d.Where(x => x.Value != null).ToDictionary(x => x.Key, x => x.Value);
            }
        }
    }
}