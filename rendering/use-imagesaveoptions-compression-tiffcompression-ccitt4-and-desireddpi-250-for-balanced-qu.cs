using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder and file.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);
        string outputPath = Path.Combine(outputFolder, "BalancedBw.tiff");

        // Create a simple source document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering with CCITT4 compression and 250 dpi.");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4, // Apply CCITT4 compression.
            Resolution = 250f                         // Desired DPI (both horizontal and vertical).
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optional: indicate success.
        Console.WriteLine($"TIFF file saved successfully to: {outputPath}");
    }
}
