using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddShapesToGroup
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to facilitate shape insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create the first child shape (a rectangle).
        Shape rect = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 150,
            Height = 100,
            Left = 0,   // Position will be relative to the group shape.
            Top = 0,
            Stroke = { Color = Color.Blue },
            FillColor = Color.LightBlue
        };

        // Create the second child shape (an ellipse).
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 100,
            Height = 100,
            Left = 50,
            Top = 30,
            Stroke = { Color = Color.Green },
            FillColor = Color.LightGreen
        };

        // Create a new GroupShape that will contain the child shapes.
        GroupShape group = new GroupShape(doc)
        {
            // Define the group's bounding box (position and size in the document).
            Bounds = new RectangleF(100, 100, 300, 200)
        };

        // Append the child shapes to the group.
        group.AppendChild(rect);
        group.AppendChild(ellipse);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeWithChildren.docx");
    }
}
