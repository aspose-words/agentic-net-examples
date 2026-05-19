using System;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Define output directory and file path
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "NoWatermark.docx");

        // Create a blank document
        Document doc = new Document();

        // (Optional) Add some content to the document
        // DocumentBuilder builder = new DocumentBuilder(doc);
        // builder.Writeln("Sample content without watermark.");

        // Save the document
        doc.Save(outputPath);

        // Validate that the document has no watermark
        bool hasNoWatermark = doc.Watermark.Type == WatermarkType.None;

        // Output validation result
        Console.WriteLine(hasNoWatermark
            ? "Validation passed: No watermark present."
            : "Validation failed: Watermark detected.");
    }
}
