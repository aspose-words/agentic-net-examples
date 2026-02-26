using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class InsertGroupShapeIntoTiff
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to add content and shapes.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;   // Position relative to the page.
        rect.Top = 50;
        rect.Stroke.Color = Color.Blue;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 100);
        ellipse.Left = 300;
        ellipse.Top = 100;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The group shape will be inserted at the current builder position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally adjust the group’s position or size.
        group.Left = 0;
        group.Top = 0;

        // Save the document as a multi‑page TIFF image.
        // Each page of the Word document becomes one frame in the TIFF file.
        doc.Save("GroupShape.tiff", SaveFormat.Tiff);
    }
}
