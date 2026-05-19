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
        builder.Writeln("Hello World!");

        // Configure image save options for TIFF with CCITT4 compression.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,
            // Render as black‑and‑white to maximize compression effectiveness.
            ImageColorMode = ImageColorMode.BlackAndWhite
        };

        // Define the output path and ensure the directory exists.
        string outputPath = Path.Combine("Output", "sample_ccitt4.tiff");
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath));

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Optionally, display the file size (no user interaction required).
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF file created at '{outputPath}' ({fileSize} bytes).");
    }
}
