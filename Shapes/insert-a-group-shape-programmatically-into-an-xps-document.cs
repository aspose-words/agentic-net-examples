using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertGroupShapeIntoXps
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);

        // Define the size and position of the group shape (in points).
        group.Width = 200;
        group.Height = 150;
        group.Left = 100;
        group.Top = 100;

        // Set the coordinate space for child shapes inside the group.
        // Here we use a 0,0 origin and a coordinate size matching the group's size.
        group.CoordOrigin = new System.Drawing.Point(0, 0);
        group.CoordSize = new System.Drawing.Size(200, 150);

        // Example child shape: a rectangle.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 80;
        rect.Left = 20;   // Position relative to the group's coordinate space.
        rect.Top = 30;
        rect.Fill.ForeColor = System.Drawing.Color.LightBlue;
        rect.Stroke.Color = System.Drawing.Color.DarkBlue;
        rect.StrokeWeight = 1.5;

        // Add the rectangle to the group.
        group.AppendChild(rect);

        // Example child shape: a line.
        Shape line = new Shape(doc, ShapeType.Line);
        line.Width = 150;
        line.Height = 0; // Height is not used for a line.
        line.Left = 10;
        line.Top = 120;
        line.Stroke.Color = System.Drawing.Color.Red;
        line.StrokeWeight = 2;
        line.Stroke.DashStyle = Aspose.Words.Drawing.DashStyle.Dash;

        // Add the line to the group.
        group.AppendChild(line);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document as XPS using default XpsSaveOptions.
        doc.Save("GroupShapeDocument.xps", new XpsSaveOptions());

        // Optionally, you can also specify additional XPS save options, e.g. high quality rendering.
        // XpsSaveOptions options = new XpsSaveOptions();
        // options.UseHighQualityRendering = true;
        // doc.Save("GroupShapeDocument_HQ.xps", options);
    }
}
