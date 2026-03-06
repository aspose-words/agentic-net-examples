using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -------------------------------------------------
        // 1. Create a group shape manually.
        // -------------------------------------------------
        GroupShape group = new GroupShape(doc);

        // Define the group’s outer bounds (position and size) in points.
        group.Bounds = new RectangleF(50, 50, 400, 400);

        // Set the internal coordinate system of the group.
        // By default it is 1000x1000; here we make it 500x500 for easier scaling.
        group.CoordSize = new Size(500, 500);
        // Move the origin to the centre of the coordinate system.
        group.CoordOrigin = new Point(-250, -250);

        // -------------------------------------------------
        // 2. Add shapes to the group.
        // -------------------------------------------------
        // Rectangle that fills the whole group coordinate space.
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = group.CoordSize.Width,
            Height = group.CoordSize.Height,
            Left = group.CoordOrigin.X,
            Top = group.CoordOrigin.Y,
            Stroke = { Color = Color.Blue, Weight = 2.0 }
        };
        group.AppendChild(rect);

        // Small red star positioned at the centre of the group.
        Shape star = new Shape(doc, ShapeType.Star)
        {
            Width = 50,
            Height = 50,
            Left = -25,   // centre (0,0) minus half width/height
            Top = -25,
            FillColor = Color.Red
        };
        group.AppendChild(star);

        // Ellipse placed near the top‑right corner of the group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 80,
            Height = 60,
            Left = 200,
            Top = -200,
            Stroke = { Color = Color.Green, Weight = 1.5 }
        };
        group.AppendChild(ellipse);

        // -------------------------------------------------
        // 3. Insert the group shape into the document.
        // -------------------------------------------------
        builder.InsertNode(group);

        // -------------------------------------------------
        // 4. Save the document.
        // -------------------------------------------------
        doc.Save("GroupShapeWithChildren.docx");
    }
}
