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
        builder.Writeln("Sample text for asymmetric DPI TIFF rendering.");

        // Configure image save options for TIFF.
        // Aspose.Words does not expose separate DpiX/DpiY properties.
        // The closest available setting is the single Resolution property,
        // which applies to both axes. Here we set it to the lower DPI (72)
        // to ensure the image is at least that resolution.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 72 // Uniform DPI; asymmetric DPI is not supported directly.
        };

        // Define output file path.
        string outputPath = "asymmetric_resolution.tiff";

        // Save the document as TIFF.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
        {
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");
        }

        // Indicate success.
        Console.WriteLine($"TIFF saved successfully at '{Path.GetFullPath(outputPath)}'.");
    }
}
