using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertGroupShapeIntoMhtml
{
    static void Main()
    {
        // Load an existing MHTML document.
        Document doc = new Document("InputDocument.mht");

        // Create a DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert two floating shapes that will be grouped.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, RelativeHorizontalPosition.Page, 100,
                                         RelativeVerticalPosition.Page, 100, 200, 150, WrapType.None);
        rect.Stroke.Color = System.Drawing.Color.Blue;

        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, RelativeHorizontalPosition.Page, 150,
                                            RelativeVerticalPosition.Page, 150, 150, 100, WrapType.None);
        ellipse.Stroke.Color = System.Drawing.Color.Green;

        // Group the two shapes. The group shape will be inserted at the current cursor position.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally adjust the group's position or size.
        group.Left = 50;
        group.Top = 50;

        // Save the modified document back to MHTML format.
        doc.Save("OutputDocument.mht", SaveFormat.Mhtml);
    }
}
