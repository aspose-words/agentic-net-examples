using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a simple 1x1 PNG image file to be used in the picture content control.
        const string imageFileName = "sample.png";
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XcZcAAAAASUVORK5CYII=");
        File.WriteAllBytes(imageFileName, pngBytes);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure there is at least one paragraph to host the content control.
        Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;

        // Create a picture content control (inline level).
        StructuredDocumentTag pictureSdt = new StructuredDocumentTag(doc, SdtType.Picture, MarkupLevel.Inline)
        {
            Title = "SamplePicture",
            Tag = "picture-tag"
        };

        // Create an image shape and embed the external PNG file.
        Shape imageShape = new Shape(doc, ShapeType.Image);
        imageShape.ImageData.SetImage(imageFileName);
        imageShape.Width = 100;   // Optional: set display size.
        imageShape.Height = 100;  // Optional: set display size.

        // Add the image shape to the picture content control.
        pictureSdt.AppendChild(imageShape);

        // Insert the picture content control into the paragraph.
        paragraph.AppendChild(pictureSdt);

        // Save the document; the image will be embedded automatically.
        doc.Save("PictureContentControl.docx");
    }
}
