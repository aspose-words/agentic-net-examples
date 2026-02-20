using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document's content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure there is a paragraph to host the group shape.
        builder.Writeln("Paragraph before the group shape.");

        // Create a new group shape associated with the document.
        GroupShape group = new GroupShape(doc);

        // Define the size and position of the group shape (in points).
        // Bounds: X, Y, Width, Height.
        group.Bounds = new RectangleF(0, 0, 200, 200);

        // Example: add a rectangle shape as a child of the group.
        Shape childRect = new Shape(doc, ShapeType.Rectangle);
        childRect.Width = 100;   // width in points
        childRect.Height = 50;   // height in points
        childRect.Left = 20;     // position relative to the group's top‑left corner
        childRect.Top = 30;
        childRect.WrapType = WrapType.None; // floating shape inside the group

        // Append the child shape to the group.
        group.AppendChild(childRect);

        // Insert the group shape into the current paragraph.
        builder.CurrentParagraph.AppendChild(group);

        // Add another paragraph after the group shape.
        builder.Writeln("Paragraph after the group shape.");

        // Save the document as a DOTX template.
        doc.Save("GroupShape.dotx", SaveFormat.Dotx);
    }
}
