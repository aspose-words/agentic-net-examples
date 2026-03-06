using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;                     // Position from the left edge of the page.
        rect.Top = 50;                      // Position from the top edge of the page.
        rect.Stroke.Color = Color.Blue;     // Outline color.

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 120;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The group’s position and size are calculated automatically.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: set group properties (e.g., make it floating and disable text wrapping).
        group.WrapType = WrapType.None;
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        group.Left = 0;
        group.Top = 0;

        // Save the document as an XPS file.
        doc.Save("GroupShape.xps", SaveFormat.Xps);
    }
}
