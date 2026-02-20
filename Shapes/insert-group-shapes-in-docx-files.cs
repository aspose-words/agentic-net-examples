using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertGroupShapeExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the cursor to a new paragraph where the group shape will be placed.
        builder.Writeln("Below is a group shape containing two rectangles:");
        builder.MoveToDocumentEnd();

        // Create a group shape. The constructor requires the owning document.
        GroupShape group = new GroupShape(doc);

        // Set the size of the group shape (in points).
        group.Width = 300;
        group.Height = 200;

        // Position the group shape on the page (floating shape).
        group.WrapType = WrapType.None;               // No text wrapping.
        group.BehindText = true;                      // Place behind text.
        group.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        group.RelativeVerticalPosition = RelativeVerticalPosition.Page;
        group.HorizontalAlignment = HorizontalAlignment.Center;
        group.VerticalAlignment = VerticalAlignment.Center;

        // Add the group shape to the document.
        // It must be inserted into a paragraph node.
        builder.CurrentParagraph.AppendChild(group);

        // ----- Add child shapes to the group -----

        // First rectangle.
        Shape rect1 = new Shape(doc, ShapeType.Rectangle);
        rect1.Width = 100;
        rect1.Height = 80;
        rect1.Left = 20;   // Position inside the group shape.
        rect1.Top = 20;
        rect1.Fill.Color = System.Drawing.Color.LightBlue;
        rect1.StrokeColor = System.Drawing.Color.DarkBlue;
        rect1.StrokeWeight = 1.5;
        group.AppendChild(rect1);

        // Second rectangle.
        Shape rect2 = new Shape(doc, ShapeType.Rectangle);
        rect2.Width = 120;
        rect2.Height = 60;
        rect2.Left = 150;
        rect2.Top = 100;
        rect2.Fill.Color = System.Drawing.Color.LightCoral;
        rect2.StrokeColor = System.Drawing.Color.Maroon;
        rect2.StrokeWeight = 1.5;
        group.AppendChild(rect2);

        // Save the document to a DOCX file.
        doc.Save("GroupShapeExample.docx");
    }
}
