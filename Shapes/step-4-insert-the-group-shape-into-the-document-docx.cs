using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) as a floating shape.
        Shape shape1 = builder.InsertShape(
            ShapeType.Rectangle,          // Shape type.
            RelativeHorizontalPosition.Page, // Horizontal reference.
            50,                           // Left position (points).
            RelativeVerticalPosition.Page,   // Vertical reference.
            50,                           // Top position (points).
            200,                          // Width (points).
            150,                          // Height (points).
            WrapType.None);               // No text wrapping.

        // Set a red stroke for the rectangle.
        shape1.Stroke.Color = Color.Red;

        // Insert the second shape (an ellipse) as a floating shape.
        Shape shape2 = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page,
            300,
            RelativeVerticalPosition.Page,
            100,
            150,
            150,
            WrapType.None);

        // Set a green stroke for the ellipse.
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes. The InsertGroupShape method automatically calculates
        // the position and size of the new GroupShape node.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Optional: modify the group shape (e.g., set a title).
        group.Title = "My Group Shape";

        // Save the document in DOCX format.
        doc.Save("GroupShapeExample.docx");
    }
}
