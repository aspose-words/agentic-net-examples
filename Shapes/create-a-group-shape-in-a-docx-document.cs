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

        // Create a rectangle shape.
        Shape rectangle = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 50,
            FillColor = Color.LightBlue,
            Stroke = { Color = Color.Black }
        };

        // Create an ellipse shape.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 80,
            Height = 80,
            FillColor = Color.LightCoral,
            Stroke = { Color = Color.DarkRed }
        };

        // Create a group shape and set its bounding box.
        GroupShape group = new GroupShape(doc);
        group.Bounds = new RectangleF(0, 0, 200, 200);

        // Add the rectangle and ellipse to the group.
        group.AppendChild(rectangle);
        group.AppendChild(ellipse);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document to a DOCX file.
        doc.Save("GroupShape.docx");
    }
}
