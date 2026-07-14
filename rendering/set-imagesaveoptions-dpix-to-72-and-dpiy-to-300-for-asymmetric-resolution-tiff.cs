using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for asymmetric TIFF resolution.");

        // Configure image save options for TIFF with different horizontal and vertical DPI.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            HorizontalResolution = 72f, // DpiX
            VerticalResolution = 300f   // DpiY
        };

        // Save the document as a TIFF image.
        string outputPath = Path.Combine(artifactsDir, "AsymmetricResolution.tiff");
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the TIFF file.");

        // Optionally, output the file size for confirmation.
        Console.WriteLine($"TIFF saved successfully. Size: {new FileInfo(outputPath).Length} bytes");
    }
}
