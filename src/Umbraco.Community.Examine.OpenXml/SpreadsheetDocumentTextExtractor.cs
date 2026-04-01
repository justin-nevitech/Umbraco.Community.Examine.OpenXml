using DocumentFormat.OpenXml.Packaging;
using System.Text;

namespace Umbraco.Community.Examine.OpenXml
{
    public class SpreadsheetDocumentTextExtractor : BaseOpenXmlTextExtractor, ISpreadsheetDocumentTextExtractor
    {
        public string GetText(Stream fileStream)
        {
            StringBuilder builder = new StringBuilder();

            using (var spreadsheetDocument = SpreadsheetDocument.Open(fileStream, false))
            {
                if (spreadsheetDocument.WorkbookPart != null)
                {
                    builder.Append(GetTextFromWorksheetParts(spreadsheetDocument.WorkbookPart));
                }
            }

            return builder.ToString().Trim();
        }
    }
}
