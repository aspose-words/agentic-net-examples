using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Load an existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) and set its appearance.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;
        shape1.Top = 20;
        shape1.Stroke.Color = System.Drawing.Color.Red;

        // Insert the second shape (an ellipse) and set its appearance.
        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = System.Drawing.Color.Green;

        // Group the two shapes into a new GroupShape node at the current cursor position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Optionally, you can add more shapes to the group after it has been created.
        // Example: clone the first shape and add it to the group.
        Shape shape3 = (Shape)shape1.Clone(true);
        group.AppendChild(shape3);

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
