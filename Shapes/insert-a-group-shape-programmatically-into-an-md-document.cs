using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a group shape that will contain other shapes.
        GroupShape group = new GroupShape(doc);

        // Set the size and position of the group shape (in points).
        group.Width = 200;
        group.Height = 150;
        group.Left = 100;   // Position from the left margin.
        group.Top = 100;    // Position from the top margin.

        // Set the group shape to be floating (not inline) and behind text.
        group.WrapType = WrapType.None;
        group.BehindText = true;
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Create a rectangle shape to add to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 80;
        rect.Height = 60;
        rect.Left = 10;   // Position inside the group coordinate space.
        rect.Top = 10;
        rect.Fill.Color = System.Drawing.Color.LightBlue;
        rect.StrokeColor = System.Drawing.Color.DarkBlue;
        rect.StrokeWeight = 1.0;

        // Add the rectangle shape as a child of the group shape.
        group.AppendChild(rect);

        // Create a second shape (ellipse) inside the group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 60;
        ellipse.Height = 60;
        ellipse.Left = 100;
        ellipse.Top = 50;
        ellipse.Fill.Color = System.Drawing.Color.LightCoral;
        ellipse.StrokeColor = System.Drawing.Color.Maroon;
        ellipse.StrokeWeight = 1.0;

        // Add the ellipse shape to the group.
        group.AppendChild(ellipse);

        // Insert the group shape into the document at the current cursor position.
        builder.InsertNode(group);

        // Save the document to a file.
        doc.Save("GroupShapeExample.docx");
    }
}
