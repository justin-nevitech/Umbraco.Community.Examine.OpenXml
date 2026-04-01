using Examine;
using Microsoft.Extensions.Logging;
using Umbraco.Cms.Core;
using Umbraco.Cms.Core.Models;
using Umbraco.Cms.Infrastructure.Examine;

namespace Umbraco.Community.Examine.OpenXml
{
    public interface IOpenXmlIndexValueSetBuilder : IValueSetBuilder<IMedia> { }

    /// <summary>
    /// Builds a ValueSet for OpenXml Documents
    /// </summary>
    public class OpenXmlIndexValueSetBuilder : IOpenXmlIndexValueSetBuilder
    {
        private OpenXmlService _openXmlService;
        private readonly ILogger<OpenXmlIndexValueSetBuilder> _logger;

        public OpenXmlIndexValueSetBuilder(OpenXmlService openXmlService, ILogger<OpenXmlIndexValueSetBuilder> logger)
        {
            _openXmlService = openXmlService;
            _logger = logger;
        }

        public IEnumerable<ValueSet> GetValueSets(params IMedia[] content)
        {
            foreach (var item in content)
            {
                var umbracoFile = item.GetValue<string>(Constants.Conventions.Media.File);
                if (string.IsNullOrWhiteSpace(umbracoFile))
                    continue;

                string fileOpenXmlContent;
                try
                {
                    fileOpenXmlContent = ExtractOpenXmlFromFile(umbracoFile);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Could not read the file {MediaFile}", umbracoFile);
                    continue;
                }
                var indexValues = new Dictionary<string, object>
                {
                    ["nodeName"] = item.Name!,
                    ["id"] = item.Id,
                    ["path"] = item.Path,
                    [OpenXmlIndexConstants.OpenXmlContentFieldName] = fileOpenXmlContent
                };

                var valueSet = new ValueSet(item.Id.ToString(), OpenXmlIndexConstants.OpenXmlCategory, item.ContentType.Alias, indexValues);

                yield return valueSet;
            }
        }

        private string ExtractOpenXmlFromFile(string filePath)
        {
            try
            {
                return _openXmlService.ExtractOpenXml(filePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Could not extract openXml from file {OpenXmlFilePath}", filePath);
                return String.Empty;
            }
        }
    }
}
