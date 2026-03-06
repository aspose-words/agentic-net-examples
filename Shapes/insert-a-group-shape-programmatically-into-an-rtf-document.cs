using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoRtf
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
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

        // Group the two shapes. The group shape is inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optional: adjust group properties (e.g., set a name or wrap type).
        group.Name = "MyGroupShape";
        group.WrapType = WrapType.None;

        // Save the document as an RTF file.
        doc.Save("GroupShapeDocument.rtf", SaveFormat.Rtf);
    }
}
