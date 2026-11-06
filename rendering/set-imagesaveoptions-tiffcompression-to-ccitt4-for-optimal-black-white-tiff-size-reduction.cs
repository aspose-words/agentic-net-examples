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

        // Configure image save options for TIFF with CCITT4 compression.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4
        };

        // Define output file path.
        string outputPath = "output.tiff";

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);
        }

        // Optional: indicate success (no user interaction required).
        Console.WriteLine("TIFF file saved successfully with CCITT4 compression.");
    }
}
