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

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a rectangle shape.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rectangle.Left = 20;                     // Position from the left edge of the page.
        rectangle.Top = 20;                      // Position from the top edge of the page.
        rectangle.Stroke.Color = Color.Red;      // Outline color.

        // Insert an ellipse shape.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The InsertGroupShape method automatically calculates
        // the position and size of the new GroupShape node.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Optional: set additional properties on the group shape.
        group.WrapType = WrapType.None;          // Make the group floating.
        group.Title = "My Group Shape";

        // Save the document to a DOCX file.
        doc.Save("GroupShapeExample.docx");
    }
}
