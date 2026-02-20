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

        // Insert a WordArt shape with custom text, font, size, and colors.
        Shape wordArt = AppendWordArt(
            doc,
            "Hello WordArt!",          // Text to display
            "Arial",                   // Font family
            400,                       // Shape width (points)
            50,                        // Shape height (points)
            Color.Yellow,              // Fill color of the shape
            Color.Blue,                // Outline (stroke) color
            ShapeType.TextPlainText);  // WordArt shape type

        // Apply additional formatting to the WordArt text.
        wordArt.TextPath.Bold = true;
        wordArt.TextPath.Italic = true;
        wordArt.TextPath.Size = 36.0; // Font size in points

        // Save the document to a DOCX file.
        doc.Save("WordArt.docx");
    }

    // Helper method that creates a WordArt shape and appends it to the document.
    private static Shape AppendWordArt(Document doc, string text, string fontFamily,
        double shapeWidth, double shapeHeight, Color fillColor, Color strokeColor, ShapeType shapeType)
    {
        // Create an inline shape of the specified WordArt type.
        Shape shape = new Shape(doc, shapeType)
        {
            WrapType = WrapType.Inline,
            Width = shapeWidth,
            Height = shapeHeight,
            FillColor = fillColor,
            StrokeColor = strokeColor
        };

        // Set the text and font for the WordArt.
        shape.TextPath.Text = text;
        shape.TextPath.FontFamily = fontFamily;

        // Append the shape inside a new paragraph at the end of the document body.
        Paragraph paragraph = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));
        paragraph.AppendChild(shape);

        return shape;
    }
}
