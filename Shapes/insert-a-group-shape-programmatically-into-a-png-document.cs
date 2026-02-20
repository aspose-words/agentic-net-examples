using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertGroupShapeIntoPng
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Define the size and position of the group on the page (in points).
        group.Bounds = new RectangleF(100, 100, 300, 200);
        // Define the internal coordinate system for child shapes.
        // CoordOrigin and CoordSize expect System.Drawing.Point and System.Drawing.Size (integer values).
        group.CoordOrigin = new Point(0, 0);
        group.CoordSize   = new Size(1000, 1000);

        // ----- Add a rectangle shape to the group -----
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 400;   // Width in points within the group's coordinate space.
        rect.Height = 200;  // Height in points within the group's coordinate space.
        rect.Left = 100;    // Position relative to the group's top‑left corner.
        rect.Top = 100;
        rect.Fill.ForeColor = Color.LightBlue;
        rect.Stroke.Color = Color.DarkBlue;
        group.AppendChild(rect);

        // ----- Add a line shape to the group -----
        Shape line = new Shape(doc, ShapeType.Line);
        line.Width = 600;
        line.Height = 0; // Height is not used for a line.
        line.Left = 200;
        line.Top = 300;
        line.Stroke.Color = Color.Red;
        line.Stroke.Weight = 2.0;
        line.Stroke.StartArrowType = ArrowType.Arrow;
        line.Stroke.EndArrowType = ArrowType.Arrow;
        group.AppendChild(line);

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // Save the document as a PNG image.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);
        doc.Save("GroupShapeOutput.png", saveOptions);
    }
}
