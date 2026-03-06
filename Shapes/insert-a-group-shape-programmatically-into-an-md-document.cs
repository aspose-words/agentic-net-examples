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

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) and set its position and stroke color.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 200, 250);
        rect.Left = 20;   // distance from the left edge of the page
        rect.Top = 20;    // distance from the top edge of the page
        rect.Stroke.Color = Color.Red;

        // Insert the second shape (an ellipse) and set its position and stroke color.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 150, 200);
        ellipse.Left = 40;
        ellipse.Top = 50;
        ellipse.Stroke.Color = Color.Green;

        // Group the two shapes into a single GroupShape node.
        // The InsertGroupShape method automatically calculates the group's bounds.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Optionally, add a third shape to the group by cloning one of the existing shapes.
        Shape clonedRect = (Shape)rect.Clone(true);
        GroupShape extendedGroup = builder.InsertGroupShape(group, clonedRect);

        // Save the document to disk.
        doc.Save("GroupShapeExample.docx");
    }
}
