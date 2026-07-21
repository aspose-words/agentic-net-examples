using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define folders for input and output.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Aspose.Words rendering example.");
        builder.Writeln("This document will be saved as a grayscale TIFF for archiving.");

        // Configure TIFF save options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the pages in grayscale.
            ImageColorMode = ImageColorMode.Grayscale,
            // Use LZW compression to keep file size reasonable.
            TiffCompression = TiffCompression.Lzw,
            // Set a typical archival resolution.
            Resolution = 150
        };

        // Save the document as a TIFF file.
        string outputPath = Path.Combine(artifactsDir, "Document.Grayscale.tiff");
        doc.Save(outputPath, tiffOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine("Grayscale TIFF saved to: " + outputPath);
    }
}
