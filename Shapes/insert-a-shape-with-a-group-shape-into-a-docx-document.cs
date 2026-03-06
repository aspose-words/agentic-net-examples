using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle shape.
        Shape shape1 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left position
            RelativeVerticalPosition.Page, 100,     // top position
            200, 150,                               // width, height
            WrapType.None);                         // no text wrapping
        shape1.Stroke.Color = Color.Red;

        // Insert a floating ellipse shape.
        Shape shape2 = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 150,
            RelativeVerticalPosition.Page, 200,
            150, 100,
            WrapType.None);
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes. The group shape is inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Adjust group properties if needed.
        group.WrapType = WrapType.None;
        group.Bounds = new RectangleF(50, 50, 300, 300);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeExample.docx");
    }
}
