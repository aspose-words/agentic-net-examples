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

        // Group the two shapes. The group’s position and size are calculated automatically.
        GroupShape group1 = builder.InsertGroupShape(rectangle, ellipse);

        // Clone the rectangle and create a second group that contains the first group and the clone.
        Shape rectangleClone = (Shape)rectangle.Clone(true);
        GroupShape group2 = builder.InsertGroupShape(group1, rectangleClone);

        // Save the resulting document.
        doc.Save("GroupShapes.docx");
    }
}
