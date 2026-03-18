using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

class ApplySemiTransparentFill
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write a line of text that we will format.
        builder.Writeln("This text has a semi‑transparent fill.");

        // Ensure the font fill is a solid fill.
        builder.Font.Fill.Solid();

        // Set the fill color (foreground color) for the text.
        builder.Font.Fill.Color = Color.Blue;

        // Set the fill transparency (0.0 = opaque, 1.0 = fully transparent).
        builder.Font.Fill.Transparency = 0.5; // 50 % transparent

        // Save the document.
        doc.Save("SemiTransparentFill.docx");
    }
}
