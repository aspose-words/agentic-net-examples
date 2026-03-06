using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create the first shape (a rectangle) and set its size, position and formatting.
        Shape rectangle = new Shape(doc, ShapeType.Rectangle)
        {
            Width = 200,
            Height = 100,
            FillColor = Color.LightBlue,
            Stroke = { Color = Color.Black }
        };
        rectangle.Left = 50;   // Position relative to the page.
        rectangle.Top = 50;

        // Create the second shape (an ellipse) and set its size, position and formatting.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse)
        {
            Width = 150,
            Height = 150,
            FillColor = Color.LightCoral,
            Stroke = { Color = Color.DarkRed }
        };
        ellipse.Left = 300;
        ellipse.Top = 80;

        // Create a GroupShape and add the two shapes as its children.
        GroupShape group = new GroupShape(doc);
        group.AppendChild(rectangle);
        group.AppendChild(ellipse);

        // Define the bounding rectangle of the group (optional, can be omitted).
        group.Bounds = new RectangleF(0, 0, 500, 300);

        // Insert the group shape into the document at the current cursor position.
        builder.InsertNode(group);

        // Save the document to a file.
        doc.Save("GroupShapeExample.docx");
    }
}
