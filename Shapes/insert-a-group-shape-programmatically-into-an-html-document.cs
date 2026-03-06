using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeIntoHtml
{
    static void Main()
    {
        // Load an existing HTML document.
        Document doc = new Document("input.html");

        // Create a DocumentBuilder to modify the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will later be grouped.
        // Shape 1: Rectangle
        Shape rect = builder.InsertShape(
            ShapeType.Rectangle,               // shape type
            RelativeHorizontalPosition.Page,   // horizontal reference
            100,                               // left position (points)
            RelativeVerticalPosition.Page,     // vertical reference
            100,                               // top position (points)
            150,                               // width (points)
            100,                               // height (points)
            WrapType.None);                    // no text wrapping

        // Optional formatting.
        rect.Stroke.Color = Color.Blue;
        rect.Fill.Color = Color.LightBlue;

        // Shape 2: Ellipse
        Shape ellipse = builder.InsertShape(
            ShapeType.Ellipse,
            RelativeHorizontalPosition.Page,
            300,
            RelativeVerticalPosition.Page,
            150,
            120,
            80,
            WrapType.None);

        ellipse.Stroke.Color = Color.Green;
        ellipse.Fill.Color = Color.LightGreen;

        // Group the two shapes. The group will be inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally adjust the group's position or size.
        // For example, move the group 50 points to the right and 30 points down.
        group.Left += 50;
        group.Top += 30;

        // Save the modified document back to HTML.
        doc.Save("output.html");
    }
}
