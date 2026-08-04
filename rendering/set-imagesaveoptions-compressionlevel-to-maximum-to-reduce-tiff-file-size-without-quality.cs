using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "Compressed.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document to demonstrate TIFF compression.");
        builder.Writeln("The file will be saved with maximum loss‑less compression.");

        // Configure ImageSaveOptions for TIFF with maximum loss‑less compression (LZW).
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Lzw; // LZW provides strong loss‑less compression.

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Optionally, output the file size.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF saved to '{outputPath}' (size: {fileSize} bytes).");
    }
}
