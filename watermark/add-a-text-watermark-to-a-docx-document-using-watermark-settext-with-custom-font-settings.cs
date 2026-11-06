using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document with sample content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");

        // Configure text watermark options with custom font settings.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 48,
            Color = Color.LightGray
            // Rotation can be set via Layout if supported; omitted here for compatibility.
        };

        // Apply the text watermark to the document.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the watermarked document.
        string outputPath = "Watermarked.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Watermarked document saved to: {outputPath}");
        }
    }
}
