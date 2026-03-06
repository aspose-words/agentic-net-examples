using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Load an existing DOCM document.
        Document doc = new Document("Input.docm");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will be grouped.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rect.Left = 50;   // Position from the left edge of the page.
        rect.Top = 50;    // Position from the top edge of the page.
        rect.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 300;
        ellipse.Top = 100;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The group shape will be inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally adjust the group shape's properties.
        group.WrapType = WrapType.None;
        group.BehindText = true;

        // Save the modified document as a DOCM file.
        doc.Save("Output.docm");
    }
}
