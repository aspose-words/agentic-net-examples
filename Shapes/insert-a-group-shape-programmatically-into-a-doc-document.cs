using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a GroupShape. The constructor requires a DocumentBase (the document itself).
        GroupShape group = new GroupShape(doc);

        // Set the size and position of the group shape (in points).
        group.Width = 300;
        group.Height = 200;
        group.Left = 100;
        group.Top = 100;

        // Set the coordinate space for child shapes inside the group.
        // Here we define a 0,0 origin and a coordinate size that matches the group's size.
        group.CoordOrigin = new System.Drawing.Point(0, 0);
        group.CoordSize = new System.Drawing.Size(300, 200);

        // Example: add a rectangle shape as a child of the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 150;
        rect.Height = 100;
        rect.Left = 50;   // Position relative to the group's coordinate space.
        rect.Top = 50;
        rect.Fill.ForeColor = System.Drawing.Color.LightBlue;
        rect.Stroke.Color = System.Drawing.Color.DarkBlue;
        rect.Stroke.Weight = 2.0;

        // Append the rectangle to the group.
        group.AppendChild(rect);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document to a DOC file.
        doc.Save("GroupShapeExample.doc");
    }
}
