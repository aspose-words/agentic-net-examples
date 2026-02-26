using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

namespace WordArtExample
{
    class Program
    {
        static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Insert a WordArt shape into the document.
            AppendWordArt(
                doc,
                "Aspose.Words WordArt",   // Text to display
                "Arial",                  // Font family
                400,                      // Shape width (points)
                50,                       // Shape height (points)
                Color.Yellow,             // Fill color of the shape
                Color.Blue,               // Outline (stroke) color
                ShapeType.TextPlainText   // WordArt shape type
            );

            // Save the document to a DOCX file.
            doc.Save("WordArt.docx");
        }

        /// <summary>
        /// Creates an inline WordArt shape, configures its text path, and adds it to the document.
        /// </summary>
        private static Shape AppendWordArt(Document doc, string text, string fontFamily,
                                          double shapeWidth, double shapeHeight,
                                          Color fillColor, Color strokeColor,
                                          ShapeType wordArtShapeType)
        {
            // Create a shape of a WordArt‑designated type.
            Shape shape = new Shape(doc, wordArtShapeType)
            {
                WrapType = WrapType.Inline,
                Width = shapeWidth,
                Height = shapeHeight,
                FillColor = fillColor,
                StrokeColor = strokeColor
            };

            // Configure the WordArt text.
            shape.TextPath.Text = text;
            shape.TextPath.FontFamily = fontFamily;
            shape.TextPath.Bold = true;   // Example formatting

            // Append the shape to a new paragraph in the document body.
            Paragraph para = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));
            para.AppendChild(shape);

            return shape;
        }
    }
}
