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
            // The shape type must be one of the Text* types to be recognized as WordArt.
            Shape wordArt = AppendWordArt(
                doc,
                "Aspose.Words WordArt!",
                "Arial",
                400,   // shape width in points
                50,    // shape height in points
                Color.LightYellow, // fill color of the shape
                Color.DarkBlue,    // outline color of the shape
                ShapeType.TextPlainText // WordArt-designated shape type
            );

            // Optionally customize the WordArt text formatting.
            wordArt.TextPath.Bold = true;
            wordArt.TextPath.Italic = true;
            wordArt.TextPath.Size = 28; // font size in points
            wordArt.TextPath.TextPathAlignment = TextPathAlignment.Center;

            // Save the document to a DOCX file.
            doc.Save("WordArtOutput.docx");
        }

        /// <summary>
        /// Creates an inline WordArt shape and appends it to a new paragraph in the document.
        /// </summary>
        private static Shape AppendWordArt(Document doc, string text, string fontFamily,
            double shapeWidth, double shapeHeight, Color fillColor, Color lineColor, ShapeType shapeType)
        {
            // Create an inline shape that will contain the WordArt.
            Shape shape = new Shape(doc, shapeType)
            {
                WrapType = WrapType.Inline,
                Width = shapeWidth,
                Height = shapeHeight,
                FillColor = fillColor,
                StrokeColor = lineColor
            };

            // Set the WordArt text and font.
            shape.TextPath.Text = text;
            shape.TextPath.FontFamily = fontFamily;

            // Append the shape to a new paragraph at the end of the document.
            Paragraph para = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));
            para.AppendChild(shape);

            return shape;
        }
    }
}
