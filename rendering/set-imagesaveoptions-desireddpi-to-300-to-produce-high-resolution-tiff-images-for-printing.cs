using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HighResolution.tiff");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some content to the document.
        builder.Writeln("This document will be saved as a high‑resolution TIFF image.");

        // Configure image save options for TIFF with 300 dpi resolution.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Sets both horizontal and vertical DPI.
            Resolution = 300
        };

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF image was not created.");

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("TIFF image saved successfully at: " + outputPath);
    }
}
