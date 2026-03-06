using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Load an existing MHTML document.
        Document doc = new Document("InputDocument.mht");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will be grouped.
        Shape rect = builder.InsertShape(
            ShapeType.Rectangle,               // shape type
            RelativeHorizontalPosition.Page,   // horizontal reference
            100,                               // left position (points)
            RelativeVerticalPosition.Page,     // vertical reference
            100,                               // top position (points)
            150,                               // width (points)
            100,                               // height (points)
            WrapType.None);                    // no text wrapping

        rect.Stroke.Color = Color.Blue;       // outline color
        rect.Fill.Color = Color.LightBlue;    // fill color

        Shape ellipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page,
            300,
            RelativeVerticalPosition.Page,
            150,
            150,
            100,
            WrapType.None);

        ellipse.Stroke.Color = Color.Green;
        ellipse.Fill.Color = Color.LightGreen;

        // Group the two shapes. The group will be inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally adjust the group's position or size.
        group.Left = 50;   // move group left
        group.Top = 50;    // move group top

        // Save the modified document back to MHTML format.
        doc.Save("OutputDocument.mht", SaveFormat.Mhtml);
    }
}
