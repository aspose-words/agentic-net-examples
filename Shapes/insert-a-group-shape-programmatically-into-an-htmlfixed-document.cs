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

        // Insert two floating shapes that will later be grouped.
        Shape rectangle = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,   // left
            RelativeVerticalPosition.Page, 100,     // top
            200, 150,                               // width, height
            WrapType.None);
        rectangle.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 150,   // left
            RelativeVerticalPosition.Page, 200,     // top
            150, 100,                               // width, height
            WrapType.None);
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The builder automatically calculates the group's bounds.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Optional: set additional properties on the group shape.
        group.Title = "SampleGroup";
        group.WrapType = WrapType.None;

        // Save the document in HTML Fixed format.
        HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions();
        doc.Save("GroupShape.html", saveOptions);
    }
}
