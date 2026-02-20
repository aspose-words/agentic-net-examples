using System;
using System.Drawing;                     // For RectangleF
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a GroupShape that will hold other shapes.
        GroupShape group = new GroupShape(doc);

        // Define the size and position of the group shape (in points).
        // Here we set a 200x200 points square at the top‑left corner of the page.
        group.Bounds = new RectangleF(0, 0, 200, 200);

        // Example: add a rectangle shape to the group.
        Shape rect = new Shape(doc, ShapeType.Rectangle);
        rect.Width = 100;
        rect.Height = 50;
        rect.Left = 20;   // Position relative to the group's coordinate space.
        rect.Top = 20;
        rect.Fill.Color = Color.LightBlue;
        group.AppendChild(rect);

        // Example: add an ellipse shape to the group.
        Shape ellipse = new Shape(doc, ShapeType.Ellipse);
        ellipse.Width = 80;
        ellipse.Height = 80;
        ellipse.Left = 100;
        ellipse.Top = 100;
        ellipse.Fill.Color = Color.LightCoral;
        group.AppendChild(ellipse);

        // Insert the group shape into the document at the current cursor position.
        // The group will be added as a floating shape.
        builder.InsertNode(group);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeInserted.docx");
    }
}
