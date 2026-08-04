using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a high‑resolution TIFF example.");

        // Configure image save options for TIFF with 600 DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set both horizontal and vertical resolution to 600 DPI.
            HorizontalResolution = 600f,
            VerticalResolution = 600f
        };

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document as a TIFF image.
        string outputPath = Path.Combine(outputDir, "HighResolution.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine("TIFF saved successfully to: " + outputPath);
    }
}
