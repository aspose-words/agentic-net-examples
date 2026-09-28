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
        builder.Writeln("Sample text for grayscale TIFF rendering.");

        // Configure TIFF save options for grayscale (8bpp indexed) output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            ImageColorMode = ImageColorMode.Grayscale, // forces 8bpp indexed grayscale
            Resolution = 300,                         // high resolution for archiving
            TiffCompression = TiffCompression.Ccitt4 // efficient CCITT Group 4 compression
        };

        // Save the document as a TIFF file.
        string outputPath = "grayscale_output.tiff";
        doc.Save(outputPath, tiffOptions);

        // Verify that the file was created and is not empty.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create TIFF file at '{outputPath}'.");

        FileInfo info = new FileInfo(outputPath);
        if (info.Length == 0)
            throw new Exception("The generated TIFF file is empty.");

        // Output a simple confirmation (non-interactive).
        Console.WriteLine($"TIFF file successfully created: {outputPath} ({info.Length} bytes)");
    }
}
