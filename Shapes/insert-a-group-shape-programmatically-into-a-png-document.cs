using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;               // Position from the left edge.
        rect.Top = 20;                // Position from the top edge.
        rect.Stroke.Color = Color.Red;

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes into a single GroupShape node.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: adjust the size of the group shape.
        group.Width = 300;
        group.Height = 300;

        // Save the document as a PNG image.
        doc.Save("GroupShape.png", SaveFormat.Png);
    }
}
