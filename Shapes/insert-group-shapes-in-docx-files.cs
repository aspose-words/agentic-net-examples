using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will later be grouped.
        Shape rectangle = builder.InsertShape(ShapeType.Rectangle, 200, 150);
        rectangle.Left = 50;   // Position from the left edge of the page.
        rectangle.Top = 50;    // Position from the top edge of the page.
        rectangle.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 150);
        ellipse.Left = 120;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The group shape will be inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rectangle, ellipse);

        // Optionally, modify the group shape (e.g., set a background fill).
        group.Fill.Color = Color.LightGray;

        // Insert another shape and group it with the previously created group shape.
        Shape triangle = new Shape(doc, ShapeType.Triangle)
        {
            Width = 100,
            Height = 100,
            Left = 200,
            Top = 200,
            Fill = { Color = Color.Yellow }
        };
        GroupShape nestedGroup = builder.InsertGroupShape(group, triangle);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeExample.docx");
    }
}
