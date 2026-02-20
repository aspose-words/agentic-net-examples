using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddShapesToGroupShape
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);
        // Set the size of the group shape (in points).
        group.Width = 300;
        group.Height = 200;
        // Define the coordinate space inside the group.
        group.CoordOrigin = new Point(0, 0);
        group.CoordSize = new Size(3000, 2000); // 10x scaling (1 point = 10 internal units).

        // Create a rectangle shape and add it to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 80;
        rect.Left = 50;   // Position inside the group coordinate space.
        rect.Top = 30;
        rect.Fill.ForeColor = Color.LightBlue;
        group.AppendChild(rect);

        // Create an ellipse shape and add it to the group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 120;
        ellipse.Height = 90;
        ellipse.Left = 180;
        ellipse.Top = 100;
        ellipse.Fill.ForeColor = Color.LightCoral;
        group.AppendChild(ellipse);

        // Insert the group shape into the document.
        builder.InsertNode(group);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeWithChildren.docx");
    }
}
