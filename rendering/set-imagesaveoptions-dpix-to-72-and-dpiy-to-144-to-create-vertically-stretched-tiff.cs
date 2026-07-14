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
        builder.Writeln("This is a sample document rendered to a vertically stretched TIFF.");

        // Configure ImageSaveOptions for TIFF with different horizontal and vertical DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.HorizontalResolution = 72f; // DpiX
        options.VerticalResolution = 144f; // DpiY

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "Stretched.tiff");

        // Save the document as a TIFF image using the specified DPI settings.
        doc.Save(outputPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Inform the user where the file was saved.
        Console.WriteLine($"TIFF saved to: {outputPath}");
    }
}
