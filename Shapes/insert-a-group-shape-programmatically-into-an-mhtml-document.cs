using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoMhtml
{
    static void Main()
    {
        // Load an existing MHTML document.
        Document doc = new Document("InputDocument.mhtml");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a new group shape and set its size.
        GroupShape group = new GroupShape(doc);
        group.Bounds = new RectangleF(0, 0, 200, 200); // Width = 200pt, Height = 200pt

        // Append the group shape to the current paragraph.
        builder.CurrentParagraph.AppendChild(group);

        // Create a child shape (e.g., a rectangle) to place inside the group.
        Shape childShape = new Shape(doc, ShapeType.Rectangle);
        childShape.Width = 100;
        childShape.Height = 100;
        childShape.Left = 50;   // Position relative to the group.
        childShape.Top = 50;
        childShape.Fill.Color = Color.LightBlue;
        childShape.StrokeColor = Color.DarkBlue;

        // Add the child shape to the group.
        group.AppendChild(childShape);

        // Save the document back to MHTML format.
        doc.Save("OutputDocument.mhtml");
    }
}
