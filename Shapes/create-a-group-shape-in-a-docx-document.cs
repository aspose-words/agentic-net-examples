using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size and position of the group shape.
        group.Width = 300;   // Width in points.
        group.Height = 200;  // Height in points.
        group.Left = 100;    // Distance from the left edge of the page.
        group.Top = 100;     // Distance from the top edge of the page.
        // Set wrapping to none so the group behaves like a floating shape.
        group.WrapType = WrapType.None;

        // Create a rectangle shape to be placed inside the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 150;
        rect.Height = 100;
        rect.Left = 0;   // Position relative to the group's coordinate space.
        rect.Top = 0;
        rect.Fill.ForeColor = Color.LightBlue;
        rect.Stroke.Color = Color.DarkBlue;
        rect.Stroke.Weight = 2;

        // Add the rectangle to the group shape.
        group.AppendChild(rect);

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeExample.docx");
    }
}
