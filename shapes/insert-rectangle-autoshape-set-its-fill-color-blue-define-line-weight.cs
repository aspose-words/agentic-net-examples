using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a floating rectangle AutoShape.
        // Position: 100 points from the left and top of the page.
        // Size: 100 points wide, 50 points high.
        // WrapType.None makes it a free‑floating shape.
        Shape rectangle = builder.InsertShape(
            ShapeType.Rectangle,
            RelativeHorizontalPosition.Page, 100,
            RelativeVerticalPosition.Page, 100,
            100, 50,
            WrapType.None);

        // Set the fill color of the rectangle to blue.
        rectangle.FillColor = Color.Blue;

        // Define the outline (stroke) weight, e.g., 3 points.
        rectangle.Stroke.Weight = 3.0;

        // Save the document to a file.
        doc.Save("RectangleShape.docx");
    }
}
