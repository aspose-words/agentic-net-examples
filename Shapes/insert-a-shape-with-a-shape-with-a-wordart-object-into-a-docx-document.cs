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

        // Insert a WordArt shape with custom text, font, size and colors.
        Shape wordArt = AppendWordArt(
            doc,
            "Aspose.Words WordArt Example",   // text
            "Calibri",                         // font family
            500,                               // shape width (points)
            100,                               // shape height (points)
            Color.LightBlue,                   // fill color
            Color.DarkBlue,                    // line (stroke) color
            ShapeType.TextPlainText);          // WordArt shape type

        // Apply additional formatting to the WordArt text.
        wordArt.TextPath.Bold = true;
        wordArt.TextPath.Italic = true;
        wordArt.TextPath.Size = 48;               // font size in points
        wordArt.TextPath.TextPathAlignment = TextPathAlignment.Center;

        // Save the document to a DOCX file.
        doc.Save("WordArtShape.docx");
    }

    /// <summary>
    /// Creates an inline WordArt shape and adds it to a new paragraph in the document.
    /// </summary>
    private static Shape AppendWordArt(Document doc, string text, string fontFamily,
        double shapeWidth, double shapeHeight, Color fill, Color stroke, ShapeType shapeType)
    {
        // Create an inline shape of the specified WordArt type.
        Shape shape = new Shape(doc, shapeType)
        {
            WrapType = WrapType.Inline,
            Width = shapeWidth,
            Height = shapeHeight,
            FillColor = fill,
            StrokeColor = stroke
        };

        // Set the WordArt text and font.
        shape.TextPath.Text = text;
        shape.TextPath.FontFamily = fontFamily;

        // Append the shape to a new paragraph.
        Paragraph para = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));
        para.AppendChild(shape);

        return shape;
    }
}
