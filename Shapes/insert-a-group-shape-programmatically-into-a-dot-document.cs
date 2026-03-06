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
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        // Insert an ellipse shape.
        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes. Position and size are calculated automatically.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Add another shape to the existing group (optional demonstration).
        Shape shape3 = (Shape)shape1.Clone(true);
        builder.InsertGroupShape(group, shape3);

        // Save the document.
        doc.Save("GroupShape.docx");
    }
}
