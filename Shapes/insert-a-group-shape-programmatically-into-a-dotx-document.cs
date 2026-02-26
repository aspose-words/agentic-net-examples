using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing DOTX template (lifecycle rule: load)
        Document doc = new Document("Template.dotx");

        // Create a DocumentBuilder to work with the document
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;               // position within the page
        rect.Top = 50;
        rect.Stroke.Color = Color.Blue;

        // Insert an ellipse shape
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 120;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes; the group size and position are calculated automatically
        // (feature rule: InsertGroupShape(params ShapeBase[]))
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: adjust group properties (e.g., make it floating and positioned at the page origin)
        group.WrapType = WrapType.None;
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        group.Left = 0;
        group.Top = 0;

        // Save the modified document as a DOTX file (lifecycle rule: save)
        doc.Save("Result.dotx");
    }
}
