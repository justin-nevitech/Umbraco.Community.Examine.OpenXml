namespace Umbraco.Community.Examine.OpenXml
{
    public interface IOpenXmlTextExtractorFactory
    {
        IOpenXmlTextExtractor GetOpenXmlTextExtractor(string extension);
    }
}