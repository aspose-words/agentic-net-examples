using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a simple paragraph so the document has visible content.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample document with a text watermark.");

        // Define custom watermark formatting.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.DarkGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = false
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the resulting document.
        string outputPath = "WatermarkedDocument.docx";
        doc.Save(outputPath);
    }
}
