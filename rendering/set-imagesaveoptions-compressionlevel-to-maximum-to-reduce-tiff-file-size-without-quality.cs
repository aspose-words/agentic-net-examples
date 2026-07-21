using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document to be saved as a TIFF image.");
        // Insert a sample image if available; otherwise just text is sufficient.
        // (No external image files are required.)

        // Configure ImageSaveOptions for TIFF with maximum loss‑less compression.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Lzw; // LZW provides lossless compression.

        // Save the document as a TIFF file.
        string tiffPath = Path.Combine(artifactsDir, "Sample.tiff");
        doc.Save(tiffPath, options);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new FileNotFoundException("TIFF file was not created.", tiffPath);

        // Optionally, output the file size for reference.
        long fileSize = new FileInfo(tiffPath).Length;
        Console.WriteLine($"TIFF file saved successfully. Size: {fileSize} bytes.");
    }
}
