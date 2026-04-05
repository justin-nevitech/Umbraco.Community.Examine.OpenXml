using Microsoft.Extensions.Logging;
using Umbraco.Cms.Core.IO;

namespace Umbraco.Community.Examine.OpenXml
{
    /// <summary>
    /// Extracts text content from OpenXml documents
    /// </summary>
    public class OpenXmlService
    {
        private readonly IOpenXmlTextExtractorFactory _openXmlTextExtractorFactory;
        private readonly MediaFileManager _mediaFileSystem;
        private readonly ILogger<OpenXmlService> _logger;

        public OpenXmlService(
            IOpenXmlTextExtractorFactory openXmlTextExtractorFactory,
            MediaFileManager mediaFileSystem,
            ILogger<OpenXmlService> logger)
        {
            _openXmlTextExtractorFactory = openXmlTextExtractorFactory;
            _mediaFileSystem = mediaFileSystem;
            _logger = logger;
        }

        /// <summary>
        /// Extract text from an OpenXml file at the given path
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string ExtractOpenXml(string filePath)
        {
            var extension = Path.GetExtension(filePath)?.TrimStart('.');

            if (string.IsNullOrEmpty(extension))
            {
                _logger.LogWarning("Unable to determine file extension for {FilePath}", filePath);
                return String.Empty;
            }

            try
            {
                using (var fileStream = _mediaFileSystem.FileSystem.OpenFile(filePath))
                {
                    if (fileStream == null)
                    {
                        _logger.LogError("Unable to open file {FilePath}", filePath);
                        return String.Empty;
                    }

                    if (fileStream.Length > OpenXmlIndexConstants.MaxFileSize)
                    {
                        _logger.LogWarning("File {FilePath} exceeds maximum size limit of {MaxSize} bytes", filePath, OpenXmlIndexConstants.MaxFileSize);
                        return String.Empty;
                    }

                    var openXmlTextExtractor = _openXmlTextExtractorFactory.GetOpenXmlTextExtractor(extension);
                    return openXmlTextExtractor.GetText(fileStream);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error extracting text from file {FilePath}", filePath);
                return String.Empty;
            }
        }
    }
}
