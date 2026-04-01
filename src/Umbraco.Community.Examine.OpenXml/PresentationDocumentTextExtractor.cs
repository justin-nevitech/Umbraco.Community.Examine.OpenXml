using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System.Text;

namespace Umbraco.Community.Examine.OpenXml
{
    public class PresentationDocumentTextExtractor : IPresentationDocumentTextExtractor
    {
        public string GetText(Stream fileStream)
        {
            StringBuilder builder = new StringBuilder();

            using (var presentationDocument = PresentationDocument.Open(fileStream, false))
            {
                if (presentationDocument.PresentationPart != null)
                {
                    if (presentationDocument.PresentationPart.SlideParts != null)
                    {
                        foreach (var slidePart in presentationDocument.PresentationPart.SlideParts)
                        {
                            AppendParagraphs(builder, slidePart.Slide);
                            AppendParagraphs(builder, slidePart.NotesSlidePart?.NotesSlide);
                        }
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
