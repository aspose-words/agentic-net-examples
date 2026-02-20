using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoEpub
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an empty paragraph – the group shape will be added to this paragraph.
        builder.Writeln();

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);

        // Define the size and position of the group shape (in points).
        // RectangleF(left, top, width, height)
        group.Bounds = new RectangleF(0, 0, 200, 200);

        // Example child shape: a rectangle.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 50;
        rect.Left = 10;   // Position inside the group.
        rect.Top = 10;
        rect.Fill.Color = Color.LightBlue;
        rect.StrokeColor = Color.DarkBlue;
        rect.StrokeWeight = 1.0;

        // Add the rectangle to the group.
        group.AppendChild(rect);

        // Example child shape: an ellipse.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 80;
        ellipse.Height = 80;
        ellipse.Left = 110;
        ellipse.Top = 60;
        ellipse.Fill.Color = Color.LightCoral;
        ellipse.StrokeColor = Color.Maroon;
        ellipse.StrokeWeight = 1.0;

        // Add the ellipse to the group.
        group.AppendChild(ellipse);

        // Append the group shape to the current paragraph.
        builder.CurrentParagraph.AppendChild(group);

        // Save the document as an EPUB file.
        doc.Save("GroupShape.epub");
    }
}
