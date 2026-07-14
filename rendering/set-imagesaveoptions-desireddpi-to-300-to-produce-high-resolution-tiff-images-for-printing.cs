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
        builder.Writeln("Sample text for high‑resolution TIFF rendering.");

        // Configure image save options for TIFF with 300 dpi.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.Resolution = 300; // Desired DPI for both horizontal and vertical resolution.

        // Define output path and ensure the directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HighResolution.tiff");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, indicate success (no interactive input required).
        Console.WriteLine("TIFF image saved successfully at: " + outputPath);
    }
}
