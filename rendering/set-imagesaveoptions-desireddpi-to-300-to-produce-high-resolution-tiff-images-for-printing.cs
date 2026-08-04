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
        builder.Writeln("This is a sample document rendered to a high‑resolution TIFF image.");

        // Configure image save options for TIFF with 300 dpi resolution.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.Resolution = 300; // Sets both horizontal and vertical DPI.

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HighResolution.tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF image.");

        // Optional: indicate success.
        Console.WriteLine($"TIFF image saved successfully at: {outputPath}");
    }
}
