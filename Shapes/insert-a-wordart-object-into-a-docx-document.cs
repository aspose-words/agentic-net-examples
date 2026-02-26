using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System.Drawing;

class WordArtExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert an inline WordArt shape. The ShapeType must start with "Text" to be a WordArt object.
        Shape wordArt = builder.InsertShape(ShapeType.TextPlainText, 480, 24);

        // Set WordArt text and formatting.
        wordArt.TextPath.Text = "Hello World! This text is bold and italic.";
        wordArt.TextPath.FontFamily = "Arial";
        wordArt.TextPath.Bold = true;
        wordArt.TextPath.Italic = true;
        wordArt.TextPath.Size = 36; // Font size in points.

        // Optional: set fill and outline colors.
        wordArt.FillColor = Color.White;
        wordArt.StrokeColor = Color.Black;

        // Save the document to a DOCX file.
        doc.Save("WordArt.docx", SaveFormat.Docx);
    }
}
