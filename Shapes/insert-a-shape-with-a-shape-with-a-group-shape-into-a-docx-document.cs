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

        // Insert a rectangle shape.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;
        rect.Top = 20;
        rect.Stroke.Color = Color.Red;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the rectangle and ellipse into a GroupShape.
        GroupShape innerGroup = builder.InsertGroupShape(rect, ellipse);

        // Clone the rectangle shape to create a new shape.
        Shape clonedRect = (Shape)rect.Clone(true);
        clonedRect.Left = 60;
        clonedRect.Top = 60;
        clonedRect.Stroke.Color = Color.Blue;

        // Group the inner group with the cloned rectangle into an outer GroupShape.
        GroupShape outerGroup = builder.InsertGroupShape(innerGroup, clonedRect);

        // Save the document containing the nested group shape.
        doc.Save("NestedGroupShape.docx");
    }
}
