using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load an existing SVG document. No explicit LoadFormat is required for SVG.
        Document doc = new Document("input.svg");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will be grouped.
        Shape rect = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left
            RelativeVerticalPosition.Page, 100,     // top
            200,                                     // width
            150,                                     // height
            WrapType.None);
        rect.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 150,
            RelativeVerticalPosition.Page, 200,
            150,
            150,
            WrapType.None);
        ellipse.Stroke.Color = Color.Green;

        // Group the shapes; the group’s position and size are calculated automatically.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally set explicit bounds for the group shape.
        group.Bounds = new RectangleF(80, 80, 300, 300);

        // Save the modified document back to SVG format.
        doc.Save("output.svg", SaveFormat.Svg);
    }
}
