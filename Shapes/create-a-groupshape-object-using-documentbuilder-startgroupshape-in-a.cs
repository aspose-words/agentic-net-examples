using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class CreateGroupShapeExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a GroupShape. This object will act as a container for other shapes.
        GroupShape group = new GroupShape(doc)
        {
            // Set the size of the group. The size should be large enough to contain all child shapes.
            Width = 250,
            Height = 120,
            // Optional: set the position of the group relative to the page.
            Left = 0,
            Top = 0,
            // Optional: make the group inline so it behaves like a regular picture.
            WrapType = WrapType.Inline
        };

        // Insert the group shape at the current cursor position.
        builder.InsertNode(group);

        // Create a rectangle shape and add it to the group.
        Shape rectangle = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 100,
            Height = 100,
            Left = 0,
            Top = 0,
            Stroke = { Color = Color.Blue }
        };
        group.AppendChild(rectangle);

        // Create an ellipse shape and add it to the group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 80,
            Height = 80,
            Left = 120,
            Top = 0,
            Stroke = { Color = Color.Green }
        };
        group.AppendChild(ellipse);

        // Save the document to a DOCX file.
        doc.Save("GroupShape.docx");
    }
}
