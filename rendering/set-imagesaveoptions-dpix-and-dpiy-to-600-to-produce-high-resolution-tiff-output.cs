using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder and file path.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "HighResolution.tiff");

        // Create a simple document with two pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a high‑resolution TIFF example.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page of the document.");

        // Configure ImageSaveOptions for TIFF with 600 DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.HorizontalResolution = 600f; // DpiX
        options.VerticalResolution = 600f;   // DpiY

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Indicate successful completion.
        Console.WriteLine($"TIFF saved to: {outputPath}");
    }
}
