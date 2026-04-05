using Examine;
using Umbraco.Cms.Core.Models;
using Umbraco.Cms.Core.Services;
using Umbraco.Cms.Infrastructure.Examine;
using Umbraco.Extensions;

namespace Umbraco.Community.Examine.OpenXml
{
    /// <summary>
    ///     Performs the data lookups required to rebuild a OpenXml index
    /// </summary>
    public class OpenXmlIndexPopulator : IndexPopulator
    {
        private static readonly string[] OpenXmlFileExtensions = new[]
        {
            OpenXmlIndexConstants.WordProcessingDocumentFileExtension,
            OpenXmlIndexConstants.PresentationDocumentFileExtension,
            OpenXmlIndexConstants.SpreadsheetDocumentFileExtension
        };

        private readonly IExamineManager _examineManager;
        private readonly IMediaService _mediaService;
        private readonly IOpenXmlIndexValueSetBuilder _mediaValueSetBuilder;
        private readonly int? _parentId;

        /// <summary>
        ///     Default constructor to lookup all content data
        /// </summary>
        /// <param name="mediaService"></param>
        /// <param name="mediaValueSetBuilder"></param>
        public OpenXmlIndexPopulator(IMediaService mediaService, IOpenXmlIndexValueSetBuilder mediaValueSetBuilder,
            IExamineManager examineManager)
            : this(null, mediaService, mediaValueSetBuilder, examineManager)
        {
        }

        /// <summary>
        ///     Optional constructor allowing specifying custom query parameters
        /// </summary>
        /// <param name="parentId"></param>
        /// <param name="mediaService"></param>
        /// <param name="mediaValueSetBuilder"></param>
        public OpenXmlIndexPopulator(int? parentId, IMediaService mediaService,
            IOpenXmlIndexValueSetBuilder mediaValueSetBuilder, IExamineManager examineManager)
        {
            _parentId = parentId;
            _mediaService = mediaService;
            _mediaValueSetBuilder = mediaValueSetBuilder;
            _examineManager = examineManager;
            RegisterIndex(OpenXmlIndexConstants.OpenXmlIndexName);
        }


        /// <summary>
        /// Remove the media from the OpenXml index
        /// </summary>
        /// <param name="media"></param>
        public void RemoveFromIndex(params IMedia[] media)
        {
            if (!_examineManager.TryGetIndex(OpenXmlIndexConstants.OpenXmlIndexName, out var index)) return;
            var ids = media.Select(m => m.Id.ToInvariantString());
            index.DeleteFromIndex(ids);
        }

        /// <summary>
        /// Remove the media from the OpenXml index
        /// </summary>
        /// <param name="mediaIds"></param>
        public void RemoveFromIndex(params int[] mediaIds)
        {
            if (!_examineManager.TryGetIndex(OpenXmlIndexConstants.OpenXmlIndexName, out var index)) return;
            var ids = mediaIds.Select(m => m.ToInvariantString());
            index.DeleteFromIndex(ids);
        }

        /// <summary>
        /// Add any media that is a openXml to the OpenXmlIndex
        /// </summary>
        /// <param name="media"></param>
        public void AddToIndex(params IMedia[] media)
        {
            if (!_examineManager.TryGetIndex(OpenXmlIndexConstants.OpenXmlIndexName, out var index)) return;
            var mediaToIndex = media.Where(m => OpenXmlFileExtensions.Contains(m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias) ?? string.Empty)).ToArray();
            if (mediaToIndex.Length > 0)
                index.IndexItems(_mediaValueSetBuilder.GetValueSets(mediaToIndex));
        }

        /// <summary>
        /// Crawl all media content and index any documents with the .openXml extension
        /// </summary>
        /// <param name="indexes"></param>
        protected override void PopulateIndexes(IReadOnlyList<IIndex> indexes)
        {
            if (indexes.Count == 0)
            {
                return;
            }

            const int pageSize = 10000;
            var pageIndex = 0;

            var mediaParentId = -1;

            if (_parentId.HasValue && _parentId.Value > 0)
            {
                mediaParentId = _parentId.Value;
            }

            long total;

            do
            {
                var media = _mediaService.GetPagedDescendants(mediaParentId, pageIndex, pageSize, out total)
                    .Where(m => OpenXmlFileExtensions.Contains(m.GetValue<string>(OpenXmlIndexConstants.UmbracoMediaExtensionPropertyAlias) ?? string.Empty))
                    .ToArray();

                if (media.Length > 0)
                {
                    foreach (var index in indexes)
                    {
                        index.IndexItems(_mediaValueSetBuilder.GetValueSets(media));
                    }
                }

                pageIndex++;
            } while (pageIndex * pageSize < total);
        }
    }
}