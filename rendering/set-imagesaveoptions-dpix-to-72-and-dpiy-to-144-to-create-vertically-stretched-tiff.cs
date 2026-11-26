using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple document with a paragraph.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document rendered to a vertically stretched TIFF.");

        // Configure image save options for TIFF with custom DPI settings.
        // Use HorizontalResolution and VerticalResolution to set different DPI values.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            HorizontalResolution = 72,   // Horizontal DPI
            VerticalResolution = 144    // Vertical DPI (stretched vertically)
        };

        // Define output file path.
        string outputPath = "stretched_output.tiff";

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
        }

        Console.WriteLine($"TIFF image saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
