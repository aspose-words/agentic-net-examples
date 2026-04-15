using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "SmallestTiff.tiff");

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello World!");

        // Configure image save options for smallest TIFF size.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt3,
            PixelFormat = ImagePixelFormat.Format1bppIndexed
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, options);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Output the resulting file size.
        Console.WriteLine($"TIFF saved successfully. Size: {new FileInfo(outputPath).Length} bytes");
    }
}
