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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Position and size of the group shape on the page (points).
        group.Left = 100;
        group.Top = 100;
        group.Width = 200;
        group.Height = 100;

        // Define the internal coordinate system for child shapes.
        // CoordOrigin is the top‑left corner of the group’s coordinate space.
        // CoordSize defines the width and height of that space.
        group.CoordOrigin = new Point(0, 0);
        group.CoordSize = new Size(2000, 1000); // 1 point = 20 units in this space.

        // Create a rectangle shape to place inside the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;   // Width in points.
        rect.Height = 50;   // Height in points.
        rect.Left = 50;     // Position relative to the group's coordinate space.
        rect.Top = 25;
        rect.Fill.ForeColor = Color.LightBlue;
        rect.Stroke.Color = Color.DarkBlue;

        // Add the rectangle as a child of the group shape.
        group.AppendChild(rect);

        // Insert the group shape into the document flow.
        builder.InsertNode(group);

        // Save the document to a file.
        doc.Save("GroupShape.docx");
    }
}
