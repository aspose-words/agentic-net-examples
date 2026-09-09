using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content.
        builder.Writeln("Aspose.Words rendering example.");
        builder.Writeln("This document will be saved as a black‑and‑white TIFF image.");

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Render the pages using 1‑bit indexed pixel format (black‑and‑white).
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Use CCITT4 compression which is suitable for 1‑bpp images.
            TiffCompression = TiffCompression.Ccitt4,
            // Optional: set a resolution (dpi) for the output image.
            Resolution = 300
        };

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Output.tiff");

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);

        // Indicate successful completion.
        Console.WriteLine($"TIFF image saved successfully to: {outputPath}");
    }
}
