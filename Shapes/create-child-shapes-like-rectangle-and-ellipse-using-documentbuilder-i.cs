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

        // Insert a rectangle shape as an inline shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 120, 60);
        rectangle.StrokeColor = Color.Blue; // Set the outline color.

        // Insert an ellipse shape as a floating shape.
        Shape ellipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page, 200,   // Left position.
            RelativeVerticalPosition.Page, 200,     // Top position.
            80, 80,                                 // Width and height.
            WrapType.None);                         // No text wrapping.
        ellipse.StrokeColor = Color.Red; // Set the outline color.

        // Group the rectangle and ellipse together.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);
        group.Left = 100; // Position the group shape.
        group.Top = 100;

        // Save the document using DML compliance to ensure proper shape rendering.
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Docx)
        {
            Compliance = OoxmlCompliance.Iso29500_2008_Transitional
        };
        doc.Save("ChildShapes.docx", saveOptions);
    }
}
