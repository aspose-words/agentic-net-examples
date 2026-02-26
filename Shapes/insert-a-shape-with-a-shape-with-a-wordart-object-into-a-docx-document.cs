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
            "Hello Aspose!",          // Text displayed by WordArt.
            "Arial",                  // Font family.
            400,                      // Width in points.
            50,                       // Height in points.
            Color.Yellow,             // Fill color of the shape.
            Color.Black,              // Outline (stroke) color.
            ShapeType.TextPlainText   // WordArt‑designated shape type.
        );

        // Optional formatting of the WordArt text.
        wordArt.TextPath.Bold = true;
        wordArt.TextPath.Italic = true;
        wordArt.TextPath.Size = 36;   // Font size in points.

        // Save the document as DOCX.
        doc.Save("WordArt.docx", SaveFormat.Docx);
    }

    // Helper method that creates an inline WordArt shape and appends it to the document.
    private static Shape AppendWordArt(Document doc, string text, string fontFamily,
        double shapeWidth, double shapeHeight, Color fillColor, Color lineColor, ShapeType wordArtShapeType)
    {
        // Create an inline Shape that will serve as a container for the WordArt.
        Shape shape = new Shape(doc, wordArtShapeType)
        {
            WrapType = WrapType.Inline,
            Width = shapeWidth,
            Height = shapeHeight,
            FillColor = fillColor,
            StrokeColor = lineColor
        };

        // Set the WordArt text and its font.
        shape.TextPath.Text = text;
        shape.TextPath.FontFamily = fontFamily;

        // Append the shape to a new paragraph at the end of the document.
        Paragraph para = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));
        para.AppendChild(shape);
        return shape;
    }
}
