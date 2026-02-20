using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline rectangle shape (width: 100 points, height: 50 points).
        builder.InsertShape(ShapeType.Rectangle, 100, 50);

        // Insert a floating rounded rectangle shape positioned relative to the page.
        // Parameters: shape type, horizontal position, left offset, vertical position, top offset,
        // width, height, and wrap type.
        builder.InsertShape(
            ShapeType.RoundRectangle,
            RelativeHorizontalPosition.Page, 150,
            RelativeVerticalPosition.Page, 200,
            120, 80,
            WrapType.None);

        // Create a line shape with custom formatting.
        Shape line = new Shape(doc, ShapeType.Line);
        line.Width = 200;                     // Length of the line.
        line.Height = 0;                      // Height is not used for a line.
        line.StrokeWeight = 3;                // Thickness of the line.
        line.Stroke.Color = Color.Blue;       // Line color.
        line.Stroke.EndCap = EndCap.Round;    // Rounded line ends.
        // Insert the line shape into the document.
        builder.InsertNode(line);

        // Save the document to a DOCX file.
        doc.Save("ShapesDocument.docx");
    }
}
