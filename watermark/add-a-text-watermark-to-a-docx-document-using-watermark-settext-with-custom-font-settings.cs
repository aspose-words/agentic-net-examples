using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add some sample text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document with a text watermark.");

        // Define custom font settings for the watermark.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 72,
            Color = Color.LightGray,
            Layout = WatermarkLayout.Diagonal
        };

        // Add the text watermark with the defined options.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Save the watermarked document.
        string outputPath = "Watermarked.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Watermarked document saved successfully to '{outputPath}'.");
        }
        else
        {
            Console.WriteLine("Failed to save the watermarked document.");
        }
    }
}
