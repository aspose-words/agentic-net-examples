using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;               // Position from the left edge.
        rect.Top = 20;                // Position from the top edge.
        rect.Stroke.Color = Color.Red;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes into a new GroupShape node at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Example: set additional properties on the group shape.
        group.WrapType = WrapType.None;
        group.ZOrder = 0;

        // Save the document to a file.
        doc.Save("GroupShapeExample.docx");
    }
}
