using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing DOTM template
        Document doc = new Document("Template.dotm");

        // Create a DocumentBuilder to work with the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new group shape and associate it with the document
        GroupShape group = new GroupShape(doc)
        {
            Width = 300,               // Width of the group (points)
            Height = 200,              // Height of the group (points)
            Left = 100,                // Distance from the left edge of the page (points)
            Top = 150,                 // Distance from the top edge of the page (points)
            CoordOrigin = new Point(0, 0),
            CoordSize = new Size(3000, 2000) // Internal coordinate system (1 point = 10 units)
        };

        // Add a simple rectangle shape inside the group
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 50,
            Left = 50,   // Position relative to the group's coordinate space
            Top = 30
        };
        rect.Fill.ForeColor = Color.LightBlue;
        rect.Stroke.Color = Color.DarkBlue;

        // Append the rectangle to the group
        group.AppendChild(rect);

        // Insert the group shape into the document at the current cursor position
        builder.InsertNode(group);

        // Save the modified document as a DOTM file (preserve macros)
        doc.Save("OutputWithGroupShape.dotm", SaveFormat.Docm);
    }
}
