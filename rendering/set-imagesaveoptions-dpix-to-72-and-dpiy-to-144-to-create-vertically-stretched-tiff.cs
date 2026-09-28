using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with one paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for vertically stretched TIFF.");

        // Configure image save options for TIFF with custom DPI.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Horizontal DPI
            HorizontalResolution = 72,
            // Vertical DPI (stretched)
            VerticalResolution = 144
        };

        // Define output file path.
        string outputPath = "stretched_output.tiff";

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the TIFF file at '{outputPath}'.");
        }

        // Indicate success.
        Console.WriteLine($"TIFF image saved successfully to '{outputPath}'.");
    }
}
