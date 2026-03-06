using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoTiff
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two primitive shapes that will be grouped.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;   // Position relative to the page.
        rect.Top = 50;
        rect.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 120;
        ellipse.Top = 100;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The group shape's position and size are calculated automatically.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally adjust the group shape (e.g., set a background fill).
        group.Fill.Color = Color.LightYellow;

        // Save the document as a TIFF image. Each page will be rendered to a separate TIFF frame.
        doc.Save("GroupShape.tiff", SaveFormat.Tiff);
    }
}
