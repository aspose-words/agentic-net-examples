using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "HighResolution.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as a high‑resolution TIFF image.");

        // Configure image save options for TIFF with 300 dpi.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.Resolution = 300; // Desired DPI for both horizontal and vertical resolution.

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Optional: indicate success (no interactive input required).
        Console.WriteLine("TIFF image saved successfully at: " + outputPath);
    }
}
