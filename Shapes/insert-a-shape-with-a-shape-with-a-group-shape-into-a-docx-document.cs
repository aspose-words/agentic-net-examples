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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (rectangle) and set its position and stroke.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = Color.Red;

        // Insert the second shape (ellipse) and set its position and stroke.
        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes. The InsertGroupShape method automatically inserts the new GroupShape
        // at the current builder position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Clone the first shape and add the clone as a child of the group.
        Shape shape3 = (Shape)shape1.Clone(true);
        group.AppendChild(shape3);

        // (Optional) Insert the group again at the current position – this demonstrates using InsertNode.
        // The group is already in the document, so this call will place a second copy.
        builder.InsertNode(group.Clone(true));

        // Save the document to a DOCX file.
        doc.Save("GroupShapeExample.docx");
    }
}
