using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as a TIFF image with maximum loss‑less compression.");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // LZW provides the strongest loss‑less compression for TIFF.
        options.TiffCompression = TiffCompression.Lzw;

        // Save the document as a TIFF file.
        string outPath = Path.Combine(artifactsDir, "Compressed.tiff");
        doc.Save(outPath, options);

        // Verify that the file was created.
        if (!File.Exists(outPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, output the file size.
        long fileSize = new FileInfo(outPath).Length;
        Console.WriteLine($"TIFF file saved successfully. Size: {fileSize} bytes.");
    }
}
