using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output file path
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output.tiff");

        // Create a simple document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering.");

        // Configure image save options for TIFF with 300 DPI.
        // In Aspose.Words the DPI is set via the Resolution property (applies to both X and Y).
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300
        };

        // Save the document as a TIFF image
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF file was not created.");
    }
}
