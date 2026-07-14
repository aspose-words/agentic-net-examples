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
        builder.Writeln("Sample text for a grayscale TIFF archive.");
        builder.Writeln("This document will be rendered using a grayscale color mode.");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render pages as grayscale images.
            ImageColorMode = ImageColorMode.Grayscale,
            // Optional: set a higher resolution for better quality.
            Resolution = 300
            // PixelFormat.Format8bppIndexed is not available in Aspose.Words' ImagePixelFormat enum.
            // The grayscale color mode provides the intended archival result.
        };

        // Save the document as a grayscale TIFF.
        string outputPath = Path.Combine(artifactsDir, "Grayscale.tiff");
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the grayscale TIFF file.");

        // Indicate successful completion.
        Console.WriteLine("Grayscale TIFF saved to: " + outputPath);
    }
}
