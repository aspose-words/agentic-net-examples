using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert initial HTML content.
        string htmlBefore = "<p>This is an HTML paragraph before the group shape.</p>";
        builder.InsertHtml(htmlBefore);

        // Insert two individual shapes that will be grouped.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rectangle.Left = 20;
        rectangle.Top = 20;
        rectangle.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 250;
        ellipse.Top = 30;
        ellipse.Stroke.Color = Color.Green;

        // Group the shapes. Position and size are calculated automatically.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Optional: adjust group properties.
        group.WrapType = WrapType.None;
        group.BehindText = true;

        // Insert additional HTML after the group shape.
        string htmlAfter = "<p>HTML after the group shape.</p>";
        builder.InsertHtml(htmlAfter);

        // Save the document as an HTML file.
        doc.Save("GroupShape.html", SaveFormat.Html);
    }
}
