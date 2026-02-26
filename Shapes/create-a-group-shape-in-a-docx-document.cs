using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create two individual shapes that will be placed inside the group.
        Shape rectangle = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 100,
            FillColor = Color.LightBlue,
            Stroke = { Color = Color.Black }
        };

        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 80,
            Height = 80,
            FillColor = Color.LightCoral,
            Stroke = { Color = Color.DarkRed }
        };

        // Create a GroupShape and add the shapes as its children.
        GroupShape group = new GroupShape(doc);
        group.AppendChild(rectangle);
        group.AppendChild(ellipse);

        // Define the outer bounds of the group shape (position and size in points).
        group.Bounds = new RectangleF(0, 0, 200, 200);

        // Optionally configure the internal coordinate system of the group.
        group.CoordSize = new Size(500, 500);      // Size of the coordinate space.
        group.CoordOrigin = new Point(-250, -250); // Move the origin to the centre.

        // Insert the group shape into the document at the current cursor position.
        builder.InsertNode(group);

        // Save the resulting document.
        doc.Save("GroupShape.docx");
    }
}
