using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a few pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for grayscale TIFF rendering.");
        for (int i = 0; i < 2; i++)
        {
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln($"Page {i + 2}");
        }

        // Define the output TIFF file path.
        string outputPath = "grayscale_output.tiff";

        // Configure ImageSaveOptions for TIFF with grayscale color mode.
        // The default pixel format for grayscale images in Aspose.Words is 8‑bpp indexed,
        // which satisfies the requirement for efficient archival TIFFs.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.Grayscale,
            Resolution = 300 // optional: set DPI for archiving quality
        };

        // Render the document to the TIFF file.
        doc.Save(outputPath, options);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");
        }

        // Output basic information about the generated file.
        Console.WriteLine($"TIFF file saved successfully: {outputPath}");
        Console.WriteLine($"File size: {new FileInfo(outputPath).Length} bytes");
    }
}
