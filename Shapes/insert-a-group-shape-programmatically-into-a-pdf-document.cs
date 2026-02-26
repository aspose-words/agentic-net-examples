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
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rectangle.Left = 20;               // Position from the left edge.
        rectangle.Top = 20;                // Position from the top edge.
        rectangle.Stroke.Color = Color.Red;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes into a single GroupShape node.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Example: set the group to have no text wrapping.
        group.WrapType = WrapType.None;

        // Save the document as a PDF file.
        doc.Save("GroupShape.pdf");
    }
}
