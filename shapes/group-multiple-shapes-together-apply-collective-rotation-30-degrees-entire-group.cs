using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

class GroupShapesWithRotation
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the first shape (a rectangle) and set its position.
        Shape rect = builder.InsertShape(ShapeType.Rectangle, 150, 100);
        rect.Left = 50;
        rect.Top = 50;
        rect.Stroke.Color = Color.Blue;
        rect.FillColor = Color.LightBlue;

        // Insert the second shape (an ellipse) and set its position.
        Shape ellipse = builder.InsertShape(ShapeType.Ellipse, 120, 120);
        ellipse.Left = 250;
        ellipse.Top = 80;
        ellipse.Stroke.Color = Color.Green;
        ellipse.FillColor = Color.LightGreen;

        // Group the two shapes together. The InsertGroupShape method automatically
        // calculates the group's bounds based on the child shapes.
        GroupShape group = builder.InsertGroupShape(rect, ellipse);

        // Apply a collective rotation of 30 degrees to the entire group.
        // Positive values rotate clockwise.
        group.Rotation = 30;

        // Save the document to a file.
        doc.Save("GroupedShapesRotated.docx");
    }
}
