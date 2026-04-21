using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a simple paragraph so the document has visible content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a text watermark.");

        // Configure watermark options: set font size and color.
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontSize = 48,          // Font size of the watermark text.
            Color = Color.Blue      // Font color of the watermark text.
            // Note: TextWatermarkOptions does not expose a spacing property;
            // spacing between repeated watermarks is controlled internally.
        };

        // Apply the text watermark with the specified options.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the resulting document.
        doc.Save("WatermarkedDocument.docx");
    }
}
