using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample text for TIFF rendering with CCITT4 compression.");

        // Configure image save options for TIFF format and set CCITT4 compression.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.TiffCompression = TiffCompression.Ccitt4;

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "SampleCcitt4.tiff");

        // Save the document as a TIFF image using the specified options.
        doc.Save(outputPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");
    }
}
