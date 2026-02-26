using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rectangle.Left = 50;   // Position from the left edge of the page.
        rectangle.Top = 50;    // Position from the top edge of the page.
        rectangle.Stroke.Color = Color.Blue;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 120;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The group’s position and size are calculated automatically.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Optional: adjust group properties (e.g., make it float behind text).
        group.WrapType = WrapType.None;
        group.BehindText = true;

        // Save the document as an EPUB file.
        doc.Save("GroupShape.epub", SaveFormat.Epub);
    }
}
