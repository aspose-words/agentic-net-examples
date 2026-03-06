using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two shapes that will be grouped.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;
        rect.Top = 20;
        rect.Stroke.Color = Color.Red;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes into a new GroupShape node.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Demonstrate nesting: clone one shape and group it with the existing group.
        Shape clonedRect = (Shape)rect.Clone(true);
        builder.InsertGroupShape(group, clonedRect);

        // Save the document.
        doc.Save("GroupShapeExample.docx");
    }
}
