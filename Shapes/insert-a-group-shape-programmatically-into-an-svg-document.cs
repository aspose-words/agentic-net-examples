using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoSvg
{
    static void Main()
    {
        // Load an existing SVG document.
        // The SVG file will be treated as a Word document containing the SVG content.
        Document doc = new Document("input.svg");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two individual shapes that will later be grouped.
        // Shape 1: Rectangle
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rect.Left = 50;   // Position relative to the page.
        rect.Top = 50;
        rect.Stroke.Color = Color.Blue;
        rect.Fill.Color = Color.LightBlue;

        // Shape 2: Ellipse
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 120, 120);
        ellipse.Left = 250;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;
        ellipse.Fill.Color = Color.LightGreen;

        // Group the two shapes into a single GroupShape node.
        // The InsertGroupShape method automatically calculates the group's position and size.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: adjust the group's bounds if a specific size is required.
        // group.Bounds = new RectangleF(40, 40, 400, 200);

        // Save the modified document back to SVG format.
        doc.Save("output.svg", SaveFormat.Svg);
    }
}
