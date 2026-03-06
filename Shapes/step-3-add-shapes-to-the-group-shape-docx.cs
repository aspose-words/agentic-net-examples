using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddShapesToGroupShape
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (rectangle) and set its position and stroke color.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;
        rect.Top = 20;
        rect.Stroke.Color = Color.Red;

        // Insert the second shape (ellipse) and set its position and stroke color.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The group shape's position and size are calculated automatically.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Insert an additional shape (triangle) that we will add to the existing group.
        Shape triangle = builder.InsertShape(ShapeType.Triangle, 100, 100);
        triangle.Left = 60;
        triangle.Top = 60;
        triangle.Stroke.Color = Color.Blue;

        // Append the new shape to the previously created group.
        group.AppendChild(triangle);

        // Save the document containing the grouped shapes.
        doc.Save("GroupShapeAdded.docx");
    }
}
