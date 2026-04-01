using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace Umbraco.Community.Examine.OpenXml
{
    public class WordProcessingDocumentTextExtractor : IWordProcessingDocumentTextExtractor
    {
        public string GetText(Stream fileStream)
        {
            StringBuilder builder = new StringBuilder();

            using (var wordprocessingDocument = WordprocessingDocument.Open(fileStream, false))
            {
                if (wordprocessingDocument.MainDocumentPart != null)
                {
                    AppendParagraphs(builder, wordprocessingDocument.MainDocumentPart.Document?.Body);

                    foreach (var headerPart in wordprocessingDocument.MainDocumentPart.HeaderParts)
                    {
                        AppendParagraphs(builder, headerPart.Header);
                    }

                    foreach (var footerPart in wordprocessingDocument.MainDocumentPart.FooterParts)
                    {
                        AppendParagraphs(builder, footerPart.Footer);
                    }

                    if (wordprocessingDocument.MainDocumentPart.FootnotesPart != null)
                    {
                        AppendParagraphs(builder, wordprocessingDocument.MainDocumentPart.FootnotesPart.Footnotes);
                    }

                    if (wordprocessingDocument.MainDocumentPart.EndnotesPart != null)
                    {
                        AppendParagraphs(builder, wordprocessingDocument.MainDocumentPart.EndnotesPart.Endnotes);
                    }
                }
            }

            return builder.ToString().Trim();
        }

        private static void AppendParagraphs(StringBuilder builder, OpenXmlCompositeElement? element)
        {
            if (element == null) return;

            foreach (var paragraph in element.Descendants<Paragraph>())
            {
                var text = paragraph.InnerText;

                if (!String.IsNullOrEmpty(text))
                {
                    builder.Append(text).Append(' ');
                }
            }
        }
    }
}
