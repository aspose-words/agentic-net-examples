using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two individual shapes that will be grouped.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes into a new GroupShape node.
        GroupShape group1 = builder.InsertGroupShape(shape1, shape2);

        // Clone one of the original shapes (deep clone) and add it to a second group together with the first group.
        Shape shape3 = (Shape)shape1.Clone(true);
        GroupShape group2 = builder.InsertGroupShape(group1, shape3);

        // Save the document to a DOCX file.
        doc.Save("Shape.CombineGroupShape.docx");
    }
}
