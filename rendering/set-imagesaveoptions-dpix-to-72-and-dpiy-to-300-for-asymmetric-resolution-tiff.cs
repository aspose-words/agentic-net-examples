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
        builder.Writeln("This is a sample document for asymmetric TIFF resolution.");

        // Ensure the output directory exists.
        string outputDir = "Output";
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "AsymmetricResolution.tiff");

        // Configure ImageSaveOptions for TIFF with different horizontal and vertical DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.HorizontalResolution = 72f; // DpiX
        options.VerticalResolution = 300f;   // DpiY

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("TIFF file was not created.");
    }
}
