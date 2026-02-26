using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing JPEG image as a Word document.
        Document doc = new Document("input.jpg");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;               // Position from the left edge of the page.
        rect.Top = 50;                // Position from the top edge of the page.
        rect.Stroke.Color = Color.Blue;
        rect.Fill.ForeColor = Color.LightBlue;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 300;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;
        ellipse.Fill.ForeColor = Color.LightGreen;

        // Group the two shapes. The group shape is inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally adjust the group's position.
        group.Left = 0;
        group.Top = 0;

        // Save the modified document back to a JPEG file.
        doc.Save("output.jpg");
    }
}
