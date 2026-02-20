using System;
using Aspose.Words;
using Aspose.Words.Drawing;

class InsertShapeExample
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to work with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a rectangle shape.
        Shape shape = new Shape(doc, ShapeType.Rectangle);

        // Set shape size (in points).
        shape.Width = 150;   // 150 points wide
        shape.Height = 80;   // 80 points high

        // Position the shape on the page (floating shape).
        shape.Left = 100;    // 100 points from the left margin
        shape.Top = 100;     // 100 points from the top margin

        // Make the shape float (not inline) and set wrapping.
        shape.WrapType = WrapType.None;          // No text wrapping
        shape.BehindText = true;                 // Place behind text
        shape.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
        shape.RelativeVerticalPosition = RelativeVerticalPosition.Page;

        // Optionally set a fill color.
        shape.Fill.ForeColor = System.Drawing.Color.LightBlue;
        shape.Fill.Visible = true;

        // Insert the shape into the document.
        builder.InsertNode(shape);

        // Save the document to a DOCX file.
        doc.Save("ShapeInsertion.docx");
    }
}
