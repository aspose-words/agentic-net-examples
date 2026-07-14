using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Build a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering.");

        // Configure image save options for TIFF with maximum loss‑less compression.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Lzw; // LZW provides lossless compression.

        // Save the document as a TIFF file.
        string outPath = Path.Combine(artifactsDir, "output.tiff");
        doc.Save(outPath, options);

        // Verify that the file was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Optionally, output the file size (no console interaction required).
        long fileSize = new FileInfo(outPath).Length;
        Console.WriteLine($"TIFF saved successfully. Size: {fileSize} bytes.");
    }
}
