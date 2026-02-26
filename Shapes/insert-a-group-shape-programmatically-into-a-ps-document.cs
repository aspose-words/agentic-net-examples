using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert shapes.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) and set its position and stroke color.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        // Insert the second shape (an ellipse) and set its position and stroke color.
        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes into a new GroupShape node.
        // The InsertGroupShape method automatically calculates the position and size of the group.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Optionally, you can add another shape to the existing group.
        Shape shape3 = (Shape)shape1.Clone(true);
        GroupShape extendedGroup = builder.InsertGroupShape(group, shape3);

        // Save the document to a file.
        doc.Save("GroupShapeExample.docx");
    }
}
