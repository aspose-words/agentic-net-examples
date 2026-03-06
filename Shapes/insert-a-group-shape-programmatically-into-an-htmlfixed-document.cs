using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoHtmlFixed
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) and set its position and stroke color.
        Shape shape1 = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        shape1.Left = 20;               // Position from the left edge of the page (points).
        shape1.Top = 20;                // Position from the top edge of the page (points).
        shape1.Stroke.Color = Color.Red;

        // Insert the second shape (an ellipse) and set its position and stroke color.
        Shape shape2 = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        shape2.Left = 40;
        shape2.Top = 50;
        shape2.Stroke.Color = Color.Green;

        // Group the two shapes into a single GroupShape node at the current cursor position.
        GroupShape group = builder.InsertGroupShape(shape1, shape2);

        // Optionally, you can add more shapes to the group after it has been created.
        // Example: clone the first shape and add it to the group.
        Shape shape3 = (Shape)shape1.Clone(true);
        group.AppendChild(shape3);

        // Save the document in HTML Fixed format. This format preserves the exact layout,
        // including floating shapes and group shapes.
        doc.Save("GroupShape.html", SaveFormat.HtmlFixed);
    }
}
