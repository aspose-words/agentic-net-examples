using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Insert a WordArt shape into the document.
        Shape wordArt = AppendWordArt(
            doc,
            "Hello Aspose!",          // Text to display.
            "Arial",                  // Font family.
            400,                      // Shape width.
            50,                       // Shape height.
            Color.Yellow,             // Fill color.
            Color.Blue,               // Outline (stroke) color.
            ShapeType.TextPlainText   // WordArt shape type.
        );

        // Apply additional formatting to the WordArt.
        wordArt.TextPath.Bold = true;
        wordArt.TextPath.Italic = true;
        wordArt.TextPath.Size = 36; // Font size in points.

        // Save the document.
        doc.Save("WordArtShape.docx");
    }

    // Helper method that creates an inline WordArt shape and adds it to a new paragraph.
    private static Shape AppendWordArt(
        Document doc,
        string text,
        string fontFamily,
        double width,
        double height,
        Color fillColor,
        Color lineColor,
        ShapeType shapeType)
    {
        // Create an inline shape of a WordArt‑designated type.
        Shape shape = new Shape(doc, shapeType)
        {
            WrapType = WrapType.Inline,
            Width = width,
            Height = height,
            FillColor = fillColor,
            StrokeColor = lineColor
        };

        // Set the WordArt text and font.
        shape.TextPath.Text = text;
        shape.TextPath.FontFamily = fontFamily;

        // Insert the shape into a new paragraph.
        Paragraph para = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));
        para.AppendChild(shape);

        return shape;
    }
}
