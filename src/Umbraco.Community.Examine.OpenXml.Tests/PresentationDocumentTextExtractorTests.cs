using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace Umbraco.Community.Examine.OpenXml.Tests;

public class PresentationDocumentTextExtractorTests
{
    private readonly PresentationDocumentTextExtractor _extractor = new();

    [Fact]
    public void GetText_WithRealPptx_ExtractsText()
    {
        using var stream = TestHelper.GetTestFileStream("test.pptx");

        var result = _extractor.GetText(stream);

        Assert.NotNull(result);
        Assert.NotEmpty(result);
    }

    [Fact]
    public void GetText_WithRealPptx_ContainsSlideContent()
    {
        using var stream = TestHelper.GetTestFileStream("test.pptx");

        var result = _extractor.GetText(stream);

        // Verify that text from slides is extracted
        Assert.NotNull(result);
        Assert.True(result.Length > 0, "Should extract text from presentation slides");
    }

    [Fact]
    public void GetText_StreamIsNotDisposedByExtractor()
    {
        using var stream = TestHelper.GetTestFileStream("test.pptx");

        _extractor.GetText(stream);

        Assert.True(stream.CanRead);
    }

    [Fact]
    public void GetText_ReturnsString_NotNull()
    {
        using var stream = TestHelper.GetTestFileStream("test.pptx");

        var result = _extractor.GetText(stream);

        Assert.IsType<string>(result);
    }

    [Fact]
    public void GetText_WithEmptyPresentation_ReturnsEmptyString()
    {
        using var stream = new MemoryStream();
        using (var doc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentationPart = doc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Equal(string.Empty, result);
    }

    [Fact]
    public void GetText_WithPresentationContainingNotes_ExtractsNotesText()
    {
        using var stream = new MemoryStream();
        using (var doc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation))
        {
            var presentationPart = doc.AddPresentationPart();
            presentationPart.Presentation = new Presentation(
                new SlideIdList(new SlideId { Id = 256, RelationshipId = "rId2" }));

            var slidePart = presentationPart.AddNewPart<SlidePart>("rId2");
            slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(),
                new Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties { Id = 2, Name = "Title" },
                        new NonVisualShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new ShapeProperties(),
                    new TextBody(
                        new D.BodyProperties(),
                        new D.Paragraph(new D.Run(new D.Text("SlideText"))))))));

            // Add notes slide
            var notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
            notesSlidePart.NotesSlide = new NotesSlide(new CommonSlideData(new ShapeTree(
                new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1, Name = "" },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()),
                new GroupShapeProperties(),
                new Shape(
                    new NonVisualShapeProperties(
                        new NonVisualDrawingProperties { Id = 2, Name = "Notes" },
                        new NonVisualShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new ShapeProperties(),
                    new TextBody(
                        new D.BodyProperties(),
                        new D.Paragraph(new D.Run(new D.Text("MyNotesText"))))))));
        }
        stream.Position = 0;

        var result = _extractor.GetText(stream);

        Assert.Contains("MyNotesText", result);
        Assert.Contains("SlideText", result);
    }
}
