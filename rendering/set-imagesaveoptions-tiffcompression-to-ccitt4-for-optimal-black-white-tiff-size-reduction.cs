using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF conversion.");

        // Configure TIFF save options: use CCITT4 compression and black‑and‑white color mode.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Ccitt4;
        options.ImageColorMode = ImageColorMode.BlackAndWhite;

        // Ensure the output folder exists.
        string outputDir = "Artifacts";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "SampleCcitt4.tiff");

        // Save the document as a TIFF file with the specified options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("TIFF file was not created.");

        // Output the file size for confirmation.
        long fileSize = new FileInfo(outputPath).Length;
        Console.WriteLine($"TIFF saved to '{outputPath}' ({fileSize} bytes) using CCITT4 compression.");
    }
}
