using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create an output directory for the generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a simple Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF conversion.");

        // Set up image save options to render the document as a TIFF
        // and apply lossless CCITT4 compression.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Ccitt4;

        // Save the document as a TIFF file using the configured options.
        string tiffPath = Path.Combine(outputDir, "SampleCcitt4.tiff");
        doc.Save(tiffPath, options);

        // Verify that the TIFF file was created successfully.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Output basic information about the saved file.
        long fileSize = new FileInfo(tiffPath).Length;
        Console.WriteLine($"TIFF saved to: {tiffPath}");
        Console.WriteLine($"File size: {fileSize} bytes");
    }
}
