using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class AddShapesToGroupShape
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will later be grouped.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.Page, 100,
                                         RelativeVerticalPosition.Page, 100, 200, 150, WrapType.None);
        rect.Stroke.Color = Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, RelativeHorizontalPosition.Page, 350,
                                            RelativeVerticalPosition.Page, 150, 150, 150, WrapType.None);
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes. The InsertGroupShape method automatically calculates the group's position and size.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Create an additional shape (a star) that will be added to the existing group.
        Shape star = new Shape(doc, ShapeType.Star)
        {
            Width = 80,
            Height = 80,
            // Position the star relative to the group's internal coordinate system.
            // Here we place it near the centre of the group.
            Left = -40,
            Top = -40,
            FillColor = Color.Yellow,
            Stroke = { Color = Color.Orange }
        };

        // Append the new shape to the group.
        group.AppendChild(star);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document.
        doc.Save("AddShapesToGroupShape.docx");
    }
}
