namespace Umbraco.Community.Examine.OpenXml
{
    public class OpenXmlTextExtractorFactory : IOpenXmlTextExtractorFactory
    {
        private readonly Dictionary<string, IOpenXmlTextExtractor> _openXmlTextExtractors;

        public OpenXmlTextExtractorFactory(IWordProcessingDocumentTextExtractor wordProcessingDocumentTextExtractor,
            IPresentationDocumentTextExtractor presentationDocumentTextExtractor,
            ISpreadsheetDocumentTextExtractor spreadsheetDocumentTextExtractor)
        {
            _openXmlTextExtractors = new Dictionary<string, IOpenXmlTextExtractor>(StringComparer.OrdinalIgnoreCase)
            {
                { OpenXmlIndexConstants.WordProcessingDocumentFileExtension, wordProcessingDocumentTextExtractor },
                { OpenXmlIndexConstants.PresentationDocumentFileExtension, presentationDocumentTextExtractor },
                { OpenXmlIndexConstants.SpreadsheetDocumentFileExtension, spreadsheetDocumentTextExtractor }
            };
        }

        public IOpenXmlTextExtractor GetOpenXmlTextExtractor(string extension)
        {
            if (_openXmlTextExtractors.TryGetValue(extension, out var openXmlTextExtractor))
            {
                return openXmlTextExtractor;
            }

            throw new NotSupportedException("No text extractor defined for the specified file type");
        }
    }
}
