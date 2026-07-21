using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class WatermarkExample
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Add some sample content so the watermark is visible.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document.");
        builder.Writeln("The watermark should appear behind this text.");

        // Configure watermark options: font size, color, and layout.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontSize = 48,               // Set the font size of the watermark.
            Color = Color.Blue,          // Set the watermark color.
            Layout = WatermarkLayout.Diagonal, // Position the watermark diagonally.
            IsSemitrasparent = false    // Make the watermark fully opaque.
        };

        // Apply the text watermark with the specified options.
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Ensure the output directory exists.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "WatermarkedDocument.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Watermarked document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the watermarked document.");
        }
    }
}
