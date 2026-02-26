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

        // Insert the first shape (a rectangle) and set its appearance.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 120, 60);
        shape1.Stroke.Color = Color.Blue;
        shape1.Fill.Color = Color.LightBlue;

        // Insert the second shape (an ellipse) and set its appearance.
        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 80, 80);
        shape2.Stroke.Color = Color.Green;
        shape2.Fill.Color = Color.LightGreen;

        // Group the two shapes together. The group shape is inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Position the group shape on the page (optional).
        group.Left = 50;
        group.Top = 100;

        // Save the document as a DOCX file.
        doc.Save("ShapeWithShape.docx");
    }
}
