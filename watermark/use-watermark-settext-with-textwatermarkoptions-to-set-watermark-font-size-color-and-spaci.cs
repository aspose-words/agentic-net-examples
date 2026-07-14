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
        builder.Writeln("Sample document content.");

        // Configure watermark options: font size, color, and layout (spacing effect).
        TextWatermarkOptions watermarkOptions = new TextWatermarkOptions
        {
            FontSize = 48,               // Set the watermark font size.
            Color = Color.Blue,          // Set the watermark color.
            FontFamily = "Arial",        // Optional: set a specific font family.
            Layout = WatermarkLayout.Diagonal // Layout influences spacing/rotation.
        };

        // Apply the text watermark with the specified options.
        doc.Watermark.SetText("CONFIDENTIAL", watermarkOptions);

        // Save the document to a file.
        string outputPath = "Watermarked.docx";
        doc.Save(outputPath);
    }
}
