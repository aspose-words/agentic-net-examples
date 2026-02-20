using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a GroupShape that will hold other shapes.
        GroupShape group = new GroupShape(doc);

        // Set the size of the group shape (in points).
        group.Width = 300;
        group.Height = 200;

        // Define the coordinate space inside the group.
        // Here we use a 10000 x 10000 coordinate system (common for DrawingML).
        group.CoordOrigin = new System.Drawing.Point(0, 0);
        group.CoordSize = new System.Drawing.Size(10000, 10000);

        // Position the group shape on the page.
        group.Left = 100;   // 100 points from the left margin
        group.Top = 100;    // 100 points from the top margin
        group.WrapType = WrapType.None; // Floating shape

        // Create a rectangle shape to place inside the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 2000;   // Width in the group's coordinate space
        rect.Height = 1000;  // Height in the group's coordinate space
        rect.Left = 1000;    // Position within the group
        rect.Top = 500;
        rect.Fill.Color = System.Drawing.Color.LightBlue;
        rect.StrokeColor = System.Drawing.Color.DarkBlue;
        rect.StrokeWeight = 2;

        // Add the rectangle to the group.
        group.AppendChild(rect);

        // Insert the group shape into the document body.
        doc.FirstSection.Body.AppendChild(group);

        // Save the document as DOCX.
        doc.Save("GroupShapeExample.docx");
    }
}
