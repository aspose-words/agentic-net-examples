using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and add some content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This document will be saved as a TIFF with asymmetric DPI.");

        // Configure ImageSaveOptions for TIFF with different horizontal and vertical DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            HorizontalResolution = 72f, // DpiX equivalent
            VerticalResolution = 300f   // DpiY equivalent
        };

        // Define output file path.
        string outputPath = "AsymmetricResolution.tiff";

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        // Optionally, inform that the process completed successfully.
        Console.WriteLine($"TIFF image saved to '{Path.GetFullPath(outputPath)}' with DPI {options.HorizontalResolution}x{options.VerticalResolution}.");
    }
}
