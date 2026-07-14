using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Build a simple document that uses a TrueType font.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Font.Name = "Arial"; // TrueType font available on most systems.
        builder.Writeln("This is a sample text to demonstrate font subsetting during rendering.");
        builder.Writeln("Rendering the document to a multi‑page TIFF with subsetted fonts reduces file size.");

        // Enable embedding of TrueType fonts and request subsetting.
        doc.FontInfos.EmbedTrueTypeFonts = true;
        doc.FontInfos.SaveSubsetFonts = true;

        // Configure TIFF rendering options.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 300,               // 300 DPI for decent quality.
            PageSet = new PageSet(0)        // Render the first page (zero‑based index).
        };

        // Save the document as a TIFF image.
        string tiffPath = Path.Combine(outputDir, "Sample.tiff");
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Output basic information about the generated file.
        FileInfo info = new FileInfo(tiffPath);
        Console.WriteLine($"TIFF saved to: {tiffPath}");
        Console.WriteLine($"File size: {info.Length} bytes");
    }
}
