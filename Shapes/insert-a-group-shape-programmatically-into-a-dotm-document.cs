using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing DOTM template.
        Document doc = new Document("Template.dotm");
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will be grouped.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;               // Position relative to the page.
        rect.Top = 50;
        rect.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 120;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;

        // Group the shapes. The group’s position and size are calculated automatically.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: set additional group properties.
        group.WrapType = WrapType.None;
        group.BehindText = true;

        // Save the document as a DOTM file.
        doc.Save("Result.dotm");
    }
}
