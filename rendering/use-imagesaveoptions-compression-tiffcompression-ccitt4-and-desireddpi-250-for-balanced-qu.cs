using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create an output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Path for the resulting TIFF file.
        string outputPath = Path.Combine(artifactsDir, "BalancedBw.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered as a balanced‑quality black‑and‑white TIFF.");

        // Configure image save options for a black‑and‑white TIFF with CCITT4 compression
        // and a resolution of 250 dpi.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4,
            ImageColorMode = ImageColorMode.BlackAndWhite,
            // The Resolution property sets both horizontal and vertical DPI.
            Resolution = 250f
        };

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        Console.WriteLine($"TIFF file successfully created at: {outputPath}");
    }
}
