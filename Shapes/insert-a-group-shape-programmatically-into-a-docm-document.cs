using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Ensure there is a paragraph to host the group shape.
        builder.Writeln("Paragraph before the group shape.");
        Paragraph paragraph = builder.CurrentParagraph;

        // Create a new GroupShape associated with the document.
        GroupShape group = new GroupShape(doc);

        // Set the size and position of the group shape (in points).
        group.Width = 200;
        group.Height = 100;
        group.Left = 50;
        group.Top = 50;

        // Define the coordinate space for child shapes inside the group.
        group.CoordOrigin = new System.Drawing.Point(0, 0);
        group.CoordSize = new System.Drawing.Size(2000, 1000); // 10x scaling (1 point = 10 units).

        // Create a rectangle shape to be a child of the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 50;
        rect.Left = 100;   // Position within the group's coordinate space.
        rect.Top = 200;
        rect.FillColor = System.Drawing.Color.LightBlue;
        rect.StrokeColor = System.Drawing.Color.DarkBlue;
        rect.StrokeWeight = 1.0;

        // Add the rectangle to the group.
        group.AppendChild(rect);

        // Optionally add more child shapes here (e.g., an ellipse, text box, etc.).

        // Insert the group shape into the document.
        paragraph.AppendChild(group);

        // Add another paragraph after the group shape.
        builder.Writeln();
        builder.Writeln("Paragraph after the group shape.");

        // Save the document as a DOCM file (macro-enabled format).
        doc.Save("GroupShapeExample.docm");
    }
}
