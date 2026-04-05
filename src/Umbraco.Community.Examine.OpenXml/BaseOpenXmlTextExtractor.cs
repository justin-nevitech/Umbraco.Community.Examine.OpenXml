using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Umbraco.Community.Examine.OpenXml
{
    public abstract class BaseOpenXmlTextExtractor
    {
        protected string GetTextFromWorksheetParts(WorkbookPart workbookPart)
        {
            StringBuilder builder = new StringBuilder();

            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
                .Elements<SharedStringItem>()
                .Take(OpenXmlIndexConstants.MaxSharedStringCount)
                .ToList();

            foreach (var worksheet in workbookPart.WorksheetParts)
            {
                if (builder.Length >= OpenXmlIndexConstants.MaxExtractedContentLength)
                    break;

                using (var reader = OpenXmlReader.Create(worksheet, false))
                {
                    while (reader.Read())
                    {
                        if (builder.Length >= OpenXmlIndexConstants.MaxExtractedContentLength)
                            break;

                        if (reader.ElementType == typeof(Cell))
                        {
                            var cell = reader.LoadCurrentElement() as Cell;

                            if (cell?.CellValue == null)
                                continue;

                            string? value;

                            if (cell.DataType != null && cell.DataType == CellValues.SharedString
                                && sharedStrings != null
                                && int.TryParse(cell.CellValue.InnerText, out var index)
                                && index >= 0 && index < sharedStrings.Count)
                            {
                                value = sharedStrings[index]?.Text?.Text;
                            }
                            else
                            {
                                value = cell.CellValue.InnerText;
                            }

                            value = value?.Trim();

                            if (!String.IsNullOrEmpty(value))
                            {
                                builder.Append(value).Append(' ');
                            }
                        }
                    }
                }
            }

            return builder.ToString();
        }
    }
}
