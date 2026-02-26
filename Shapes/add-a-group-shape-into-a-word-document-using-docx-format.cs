using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will later be grouped
        Shape shape1 = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            150, 100,
            WrapType.None);
        shape1.Stroke.Color = Color.Blue;

        Shape shape2 = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 200,
            RelativeVerticalPosition.Page, 200,
            150, 100,
            WrapType.None);
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes into a new GroupShape node and insert it at the current cursor position
        GroupShape group = builder.InsertGroupShape(new Shape[] { shape1, shape2 });

        // Optionally adjust the group's properties (size, position, etc.)
        group.Bounds = new RectangleF(80, 80, 300, 300);
        group.CoordSize = new Size(1000, 1000);

        // Save the document in DOCX format
        doc.Save("GroupShapeExample.docx");
    }
}
