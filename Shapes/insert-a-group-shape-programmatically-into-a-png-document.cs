using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;               // Position from the left edge of the page.
        rect.Top = 50;                // Position from the top edge of the page.
        rect.Stroke.Color = Color.Blue;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 100;
        ellipse.Top = 100;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes into a single GroupShape node.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally adjust the group's bounding box.
        group.Bounds = new RectangleF(0, 0, 300, 300);

        // Save the document as a PNG image (each page becomes a separate PNG file).
        doc.Save("GroupShape.png", SaveFormat.Png);
    }
}
