using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document (will be saved as DOTM later)
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape (floating) at a specific position
        Shape rect = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left
            RelativeVerticalPosition.Page, 100,     // top
            200,                                     // width
            150,                                     // height
            WrapType.None);
        rect.Stroke.Color = Color.Blue;

        // Insert an ellipse shape (floating) at a specific position
        Shape ellipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 150,   // left
            RelativeVerticalPosition.Page, 200,     // top
            150,                                     // width
            100,                                     // height
            WrapType.None);
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The group shape is inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: adjust group properties (size, position, wrapping)
        group.WrapType = WrapType.None;
        group.Bounds = new RectangleF(80, 80, 300, 300);

        // Save the document as a macro‑enabled template (DOTM)
        doc.Save("GroupShape.dotm");
    }
}
