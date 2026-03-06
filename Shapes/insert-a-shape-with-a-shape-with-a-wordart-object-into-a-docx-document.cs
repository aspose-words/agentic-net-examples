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

        // Initialize a DocumentBuilder for convenient editing.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Append a new paragraph that will contain the WordArt shape.
        Paragraph paragraph = (Paragraph)doc.FirstSection.Body.AppendChild(new Paragraph(doc));

        // Create a WordArt shape. The ShapeType that starts with "Text" designates a WordArt object.
        Shape wordArt = new Shape(doc, ShapeType.TextPlainText)
        {
            // Make the shape inline so it behaves like a character in the paragraph.
            WrapType = WrapType.Inline,
            // Define the size of the shape in points.
            Width = 400,
            Height = 50,
            // Set fill and outline colors.
            FillColor = Color.Yellow,
            StrokeColor = Color.Red
        };

        // Configure the text displayed by the WordArt.
        wordArt.TextPath.Text = "Aspose.Words WordArt";
        wordArt.TextPath.FontFamily = "Calibri";
        wordArt.TextPath.Size = 36;          // Font size in points.
        wordArt.TextPath.Bold = true;        // Bold formatting.
        wordArt.TextPath.Italic = true;      // Italic formatting.

        // Insert the shape into the paragraph.
        paragraph.AppendChild(wordArt);

        // Save the document to a DOCX file.
        doc.Save("WordArt.docx");
    }
}
