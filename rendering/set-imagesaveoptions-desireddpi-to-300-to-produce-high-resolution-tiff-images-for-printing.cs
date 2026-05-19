using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a high‑resolution TIFF rendering example.");

        // Configure image save options for TIFF with 300 DPI.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300 // Sets both horizontal and vertical DPI.
        };

        // Save the document as a TIFF image.
        string tiffPath = Path.Combine(outputDir, "HighResolution.tiff");
        doc.Save(tiffPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Optionally, report success.
        Console.WriteLine($"TIFF image saved successfully at: {tiffPath}");
    }
}
