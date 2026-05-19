using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "DocumentWithCcitt4Compression.tiff");

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample content.
        builder.Writeln("This is a sample document rendered to TIFF with CCITT4 compression.");
        builder.Writeln("Aspose.Words rendering example.");

        // Configure image save options for TIFF with CCITT4 compression.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            TiffCompression = TiffCompression.Ccitt4
        };

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);

        // Optionally, report success.
        Console.WriteLine($"TIFF file saved successfully to: {outputPath}");
    }
}
