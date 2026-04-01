namespace Umbraco.Community.Examine.OpenXml
{
    public interface IOpenXmlTextExtractor
    {
        string GetText(Stream fileStream);
    }
}