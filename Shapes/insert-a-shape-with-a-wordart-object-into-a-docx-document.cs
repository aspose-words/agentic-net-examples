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

        // Insert a WordArt shape into the document.
        // The shape type must start with "Text" to be recognized as WordArt.
        Shape wordArt = AppendWordArt(
            doc,
            "Hello Aspose.Words WordArt!",
            "Arial",
            400,   // shape width in points
            50,    // shape height in points
            Color.Yellow,   // fill color of the shape
            Color.Black,    // outline (stroke) color
            ShapeType.TextPlainText);

        // Example of setting additional WordArt formatting.
        wordArt.TextPath.Bold = true;
        wordArt.TextPath.Italic = true;
        wordArt.TextPath.Size = 28; // font size in points
        wordArt.TextPath.TextPathAlignment = TextPathAlignment.Center;

        // Verify that the shape is indeed a WordArt object.
        Console.WriteLine("IsWordArt: " + wordArt.IsWordArt);

        // Save the document to a DOCX file.
        doc.Save("WordArtExample.docx");
    }

    /// <summary>
    /// Creates an inline WordArt shape and appends it to a new paragraph in the document.
    /// </summary>
    private static Shape AppendWordArt(
        Document doc,
        string text,
        string textFontFamily,
        double shapeWidth,
        double shapeHeight,
        Color fillColor,
        Color lineColor,
        ShapeType wordArtShapeType)
    {
        // Create a shape with a WordArt‑designated ShapeType.
        Shape shape = new Shape(doc, wordArtShapeType)
        {
            WrapType = WrapType.Inline,
            Width = shapeWidth,
            Height = shapeHeight,
            FillColor = fillColor,
            StrokeColor = lineColor
        };

        // Set the WordArt text and font.
        shape.TextPath.Text = text;
        shape.TextPath.FontFamily = textFontFamily;

        // Append the shape to a new paragraph at the end of the document.
        Paragraph para = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));
        para.AppendChild(shape);

        return shape;
    }
}
