using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Fonts;
using Aspose.Words.Saving;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple DOCX document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document for TIFF rendering.");
        builder.Writeln("It contains a few lines of text to demonstrate the output.");

        // Configure FontSettings (without using the prohibited EnableOpenTypeFeatures property).
        FontSettings fontSettings = new FontSettings();
        doc.FontSettings = fontSettings;

        // Prepare image save options for 1bpp TIFF with CCITT compression.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Use CCITT Group 3 compression which works well with 1‑bit images.
            TiffCompression = TiffCompression.Ccitt3,
            // Set pixel format to 1‑bit indexed to achieve minimal file size.
            PixelFormat = ImagePixelFormat.Format1bppIndexed,
            // Optional: set a reasonable resolution.
            Resolution = 300
        };

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RenderedOutput.tiff");

        // Save the document as a TIFF image.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The TIFF file was not created.", outputPath);

        // Output the file size for informational purposes.
        FileInfo info = new FileInfo(outputPath);
        Console.WriteLine($"TIFF file created: {outputPath}");
        Console.WriteLine($"File size: {info.Length} bytes");
    }
}
