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

        // Configure image save options for TIFF with 300 DPI.
        // In Aspose.Words the DPI is set via the Resolution property (applies to both axes).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.Resolution = 300; // Sets both DpiX and DpiY to 300.

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("TIFF file was not created.", outputPath);
        }

        // Output a confirmation (no user interaction required).
        Console.WriteLine($"TIFF file saved successfully at '{Path.GetFullPath(outputPath)}' with 300 DPI.");
    }
}
