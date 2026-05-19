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

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as a grayscale TIFF for archiving.");
        builder.Writeln("The image is rendered using an indexed pixel format to keep the file size low.");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the pages in grayscale.
            ImageColorMode = ImageColorMode.Grayscale,
            // Use an indexed pixel format (1‑bit) which is the closest available option.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Apply CCITT Group 4 compression – efficient for bi‑level images.
            TiffCompression = TiffCompression.Ccitt4,
            // Set a reasonable resolution for archival quality.
            Resolution = 300
        };

        // Save the document as a TIFF file.
        string outputPath = Path.Combine(artifactsDir, "GrayscaleIndexed.tiff");
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Optionally, report the file size (useful for archiving considerations).
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF saved successfully: {outputPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}
