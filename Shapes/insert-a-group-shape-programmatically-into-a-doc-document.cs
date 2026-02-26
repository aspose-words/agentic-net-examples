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

        // Initialize a DocumentBuilder for the document.
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

        // Group the two shapes into a new GroupShape node and insert it at the current cursor position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Optionally, adjust the group's properties (e.g., set a name or change wrapping).
        group.Name = "MyGroupShape";
        group.WrapType = WrapType.None;

        // Save the document to a DOC file.
        doc.Save("GroupShape.doc");
    }
}
