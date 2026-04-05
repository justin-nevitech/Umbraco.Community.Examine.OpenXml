namespace Umbraco.Community.Examine.OpenXml
{
    public static class OpenXmlIndexConstants
    {
        public const string OpenXmlIndexName = "OpenXmlIndex";
        public const string OpenXmlContentFieldName = "fileTextContent";
        public const string UmbracoMediaExtensionPropertyAlias = "umbracoExtension";
        public const string WordProcessingDocumentFileExtension = "docx";
        public const string PresentationDocumentFileExtension = "pptx";
        public const string SpreadsheetDocumentFileExtension = "xlsx";
        public const string OpenXmlCategory = "openxml";
        public const int MaxExtractedContentLength = 10 * 1024 * 1024; // 10MB
        public const long MaxFileSize = 100 * 1024 * 1024; // 100MB
        public const int MaxSharedStringCount = 1_000_000;
        public const long MaxCharactersInPart = 10_000_000;
    }
}