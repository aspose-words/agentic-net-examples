using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoJpegDocument
{
    static void Main()
    {
        // Path to the JPEG image that will be inserted into the document.
        string jpegPath = @"C:\Images\sample.jpg";

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the JPEG image as a floating shape.
        // The returned Shape object can be further manipulated (size, position, etc.).
        Shape imageShape = builder.InsertImage(jpegPath);
        imageShape.WrapType = WrapType.None;               // No text wrapping.
        imageShape.BehindText = true;                      // Place behind text.
        imageShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        imageShape.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        imageShape.HorizontalAlignment = HorizontalAlignment.Center;
        imageShape.VerticalAlignment = VerticalAlignment.Center;

        // Insert an additional shape (e.g., a rectangle) that will be grouped with the image.
        Shape rectShape = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rectShape.Left = 100;      // Position relative to the page.
        rectShape.Top = 100;
        rectShape.Stroke.Color = Color.Red;

        // Group the two shapes into a single GroupShape node.
        // The InsertGroupShape method automatically calculates the group's position and size.
        GroupShape group = builder.InsertGroupShape(imageShape, rectShape);

        // Optionally adjust the group's properties (e.g., move the group).
        group.Left = 50;
        group.Top = 50;

        // Save the resulting document. The document can be saved in any supported format (e.g., DOCX).
        doc.Save(@"C:\Output\GroupedImageDocument.docx");
    }
}
