using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an empty paragraph to host the group shape
        builder.Writeln();

        // Create a group shape and add it to the document
        GroupShape group = new GroupShape(doc);
        group.Width = 300;               // Width of the group shape in points
        group.Height = 200;              // Height of the group shape in points
        group.CoordSize = new Size(1000, 1000); // Coordinate space for child shapes
        group.CoordOrigin = new Point(0, 0);

        // Insert the group shape into the document after the current paragraph
        builder.CurrentParagraph.ParentNode.InsertAfter(group, builder.CurrentParagraph);

        // -------------------------------------------------
        // Add a rectangle shape to the group shape
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 150;
        rect.Height = 100;
        rect.Left = 100;   // Position inside the group coordinate space
        rect.Top = 50;
        rect.Fill.ForeColor = Color.LightBlue;
        rect.Fill.Visible = true;

        // Add a line shape to the group shape
        Shape line = new Shape(doc, ShapeType.Line);
        line.Width = 200;
        line.Height = 0;   // Height is not used for a line
        line.Left = 50;
        line.Top = 150;
        line.StrokeColor = Color.DarkRed;
        line.StrokeWeight = 2.0;
        line.Stroke.DashStyle = DashStyle.Dash;

        // Append the child shapes to the group shape
        group.AppendChild(rect);
        group.AppendChild(line);

        // -------------------------------------------------
        // Save the document to a DOCX file
        doc.Save("GroupShapeWithChildren.docx");
    }
}
