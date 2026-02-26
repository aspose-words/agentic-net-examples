using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) as a floating shape.
        Shape rect = builder.InsertShape(
            ShapeType.Rectangle,          // Shape type.
            RelativeHorizontalPosition.Page, // Horizontal reference.
            100,                          // Left position (points).
            RelativeVerticalPosition.Page,   // Vertical reference.
            100,                          // Top position (points).
            200,                          // Width (points).
            150,                          // Height (points).
            WrapType.None);               // No text wrapping.

        // Set a red stroke for the rectangle.
        rect.Stroke.Color = Color.Red;

        // Insert the second shape (an ellipse) as a floating shape.
        Shape ellipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page,
            350,
            RelativeVerticalPosition.Page,
            150,
            150,
            150,
            WrapType.None);

        // Set a green stroke for the ellipse.
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The InsertGroupShape method automatically calculates
        // the position and size of the new GroupShape based on the supplied shapes.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: modify the group shape (e.g., set a title).
        group.Title = "My Group Shape";

        // Save the document to a DOCX file.
        doc.Save("GroupShapeExample.docx");
    }
}
