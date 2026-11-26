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
        builder.Writeln("Sample text for asymmetric DPI TIFF rendering.");

        // Configure ImageSaveOptions for TIFF with asymmetric DPI.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Horizontal (X) resolution.
            HorizontalResolution = 72,
            // Vertical (Y) resolution.
            VerticalResolution = 300
        };

        // Define output file path.
        string outputPath = "asymmetric_dpi.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
        }

        // Output a confirmation.
        Console.WriteLine($"TIFF file saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
