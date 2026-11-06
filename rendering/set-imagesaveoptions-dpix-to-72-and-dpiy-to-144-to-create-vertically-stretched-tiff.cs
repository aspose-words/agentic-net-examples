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
        builder.Writeln("Sample text for TIFF rendering.");

        // Define the output TIFF file path.
        string tiffPath = "Output.tiff";

        // Configure image save options with custom DPI values.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Horizontal DPI (X axis)
            HorizontalResolution = 72,
            // Vertical DPI (Y axis) – this will stretch the image vertically.
            VerticalResolution = 144
        };

        // Render the document to a TIFF image using the specified options.
        doc.Save(tiffPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
        {
            throw new Exception("Failed to create the TIFF file.");
        }

        // Indicate success.
        Console.WriteLine("TIFF file created successfully at: " + Path.GetFullPath(tiffPath));
    }
}
