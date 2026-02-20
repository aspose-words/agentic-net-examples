using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial text.
        builder.Writeln("Paragraph before the group shape.");

        // Create a group shape and set its size, position and wrapping.
        GroupShape group = new GroupShape(doc);
        group.Width = 200;               // Width in points.
        group.Height = 100;              // Height in points.
        group.Left = 100;                // Horizontal position.
        group.Top = 100;                 // Vertical position.
        group.WrapType = WrapType.None;  // Floating shape.
        group.BehindText = true;         // Place behind text.

        // Insert the group shape into the current paragraph.
        builder.CurrentParagraph.AppendChild(group);

        // ----- Add a rectangle shape inside the group -----
        Shape rectangle = new Shape(doc, ShapeType.Rectangle);
        rectangle.Width = 80;
        rectangle.Height = 60;
        rectangle.Left = 10;   // Position relative to the group.
        rectangle.Top = 10;
        rectangle.FillColor = Color.LightBlue;
        rectangle.StrokeColor = Color.DarkBlue;
        rectangle.StrokeWeight = 1.0;
        group.AppendChild(rectangle);

        // ----- Add an ellipse shape inside the group -----
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 60;
        ellipse.Height = 60;
        ellipse.Left = 100;
        ellipse.Top = 20;
        ellipse.FillColor = Color.LightCoral;
        ellipse.StrokeColor = Color.Maroon;
        ellipse.StrokeWeight = 1.0;
        group.AppendChild(ellipse);

        // Save the document to a DOCX file.
        doc.Save("GroupShape.docx");
    }
}
