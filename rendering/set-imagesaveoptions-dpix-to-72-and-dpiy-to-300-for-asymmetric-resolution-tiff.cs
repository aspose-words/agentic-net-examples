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
        builder.Writeln("This document will be saved as a TIFF with asymmetric DPI.");

        // Configure image save options for TIFF.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Set horizontal DPI to 72 and vertical DPI to 300.
            HorizontalResolution = 72f,
            VerticalResolution = 300f
        };

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AsymmetricResolution.tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, output the file size for confirmation.
        Console.WriteLine($"TIFF saved successfully. Size: {new FileInfo(outputPath).Length} bytes");
    }
}
