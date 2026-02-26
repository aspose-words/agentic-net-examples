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

        // Insert a rectangle shape.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 20;               // Position relative to the page.
        rect.Top = 20;
        rect.Stroke.Color = Color.Blue;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 100);
        ellipse.Left = 250;
        ellipse.Top = 30;
        ellipse.Stroke.Color = Color.Green;

        // Group the rectangle and ellipse into a single GroupShape.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Create a star shape that will be added directly to the group.
        Shape star = new Shape(doc, ShapeType.Star)
        {
            Width = 80,
            Height = 80,
            // Position inside the group's own coordinate system.
            Left = -40,
            Top = -40,
            FillColor = Color.Yellow
        };

        // Append the star to the existing group.
        group.AppendChild(star);

        // Adjust the group's internal coordinate system (optional).
        group.CoordSize = new Size(500, 500);   // Scale of the group's coordinate plane.
        group.CoordOrigin = new Point(-250, -250); // Move origin to the centre of the group.

        // Save the document containing the complex grouped shapes.
        doc.Save("ComplexGroupShape.docx");
    }
}
