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
        builder.Writeln("Sample text for asymmetric DPI TIFF.");

        // Set up ImageSaveOptions for TIFF with asymmetric DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.HorizontalResolution = 72f; // DpiX
        options.VerticalResolution = 300f; // DpiY

        // Save the document as a TIFF file.
        string outputPath = "AsymmetricDpi.tiff";
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the TIFF file.");
    }
}
