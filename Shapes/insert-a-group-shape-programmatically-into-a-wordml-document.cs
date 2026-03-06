using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape and set its visual properties.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;
        rect.Top = 20;
        rect.Stroke.Color = Color.Red;

        // Insert an ellipse shape and set its visual properties.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the rectangle and ellipse. The group shape is inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Demonstrate grouping a group shape with another shape (clone of the rectangle).
        Shape clonedRect = (Shape)rect.Clone(true);
        builder.InsertGroupShape(group, clonedRect);

        // Save the document containing the group shape.
        doc.Save("GroupShape.docx");
    }
}
