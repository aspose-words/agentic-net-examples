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

        // Insert two floating shapes that will be grouped.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rectangle.Left = 50;               // Position from the left edge of the page.
        rectangle.Top = 50;                // Position from the top edge of the page.
        rectangle.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 120;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The group shape is inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Optional: adjust group properties (e.g., make it float behind text).
        group.WrapType = WrapType.None;
        group.BehindText = true;

        // Save the document as a Markdown file.
        doc.Save("GroupShape.md", SaveFormat.Markdown);
    }
}
