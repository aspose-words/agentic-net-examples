using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a GroupShape that will hold other shapes.
        GroupShape group = new GroupShape(doc);

        // Define the size and position of the group shape (in points).
        // Here we set a 200x200 points square positioned at (0,0) relative to its anchor.
        group.Bounds = new RectangleF(0, 0, 200, 200);

        // Optional: set wrapping and positioning properties.
        group.WrapType = WrapType.None;               // Floating shape.
        group.BehindText = true;                      // Place behind text.
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        group.HorizontalAlignment = HorizontalAlignment.Center;
        group.VerticalAlignment = VerticalAlignment.Center;

        // Create a rectangle shape to add to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;      // Width in points.
        rect.Height = 50;      // Height in points.
        rect.Left = 20;        // Position inside the group.
        rect.Top = 30;
        rect.Fill.Color = Color.Yellow;
        rect.StrokeColor = Color.Black;
        rect.StrokeWeight = 0.5;

        // Append the rectangle shape to the group.
        group.AppendChild(rect);

        // Insert the group shape into the document at the current builder position.
        builder.InsertNode(group);

        // Save the document as DOCX.
        doc.Save("GroupShape.docx", SaveFormat.Docx);
    }
}
