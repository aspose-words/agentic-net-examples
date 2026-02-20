using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder for convenient insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);

        // Set the size of the group shape (in points).
        group.Width = 300;
        group.Height = 200;

        // Define the coordinate space for child shapes.
        // CoordOrigin is the top‑left corner of the group’s internal canvas.
        group.CoordOrigin = new System.Drawing.Point(0, 0);
        // CoordSize defines the width and height of the internal canvas.
        group.CoordSize = new System.Drawing.Size(3000, 2000); // 10× scaling (points * 10)

        // Position the group shape on the page.
        group.Left = 100;
        group.Top = 100;
        group.WrapType = WrapType.None; // Floating shape.
        group.BehindText = true;

        // Create a rectangle shape to add to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 1500;   // 150 points (since CoordSize is scaled by 10)
        rect.Height = 1000;  // 100 points
        rect.Left = 0;       // Position within the group’s canvas.
        rect.Top = 0;
        rect.Fill.ForeColor = System.Drawing.Color.LightBlue;
        rect.Stroke.Color = System.Drawing.Color.DarkBlue;

        // Add the rectangle to the group.
        group.AppendChild(rect);

        // Create a second shape (e.g., an ellipse) inside the group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 1000;
        ellipse.Height = 800;
        ellipse.Left = 1500; // Position next to the rectangle.
        ellipse.Top = 500;
        ellipse.Fill.ForeColor = System.Drawing.Color.LightCoral;
        ellipse.Stroke.Color = System.Drawing.Color.Maroon;

        group.AppendChild(ellipse);

        // Insert the group shape into the document.
        // Here we insert it after the current paragraph.
        builder.InsertNode(group);

        // Save the document to a file.
        doc.Save("GroupShapeExample.docx");
    }
}
