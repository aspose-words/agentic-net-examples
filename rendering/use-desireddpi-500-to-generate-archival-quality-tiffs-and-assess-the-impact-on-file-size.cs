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
        builder.Writeln("This is a sample document for TIFF rendering with high DPI.");

        // Define output file paths.
        string defaultTiffPath = Path.Combine(Directory.GetCurrentDirectory(), "output_default.tiff");
        string highDpiTiffPath = Path.Combine(Directory.GetCurrentDirectory(), "output_high.tiff");

        // Render TIFF with default DPI.
        doc.Save(defaultTiffPath, SaveFormat.Tiff);

        // Render TIFF with high DPI (500).
        ImageSaveOptions highDpiOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // The Resolution property sets the DPI for image rendering.
            Resolution = 500
        };
        doc.Save(highDpiTiffPath, highDpiOptions);

        // Verify that both files were created.
        if (!File.Exists(defaultTiffPath))
            throw new FileNotFoundException("Default DPI TIFF was not created.", defaultTiffPath);
        if (!File.Exists(highDpiTiffPath))
            throw new FileNotFoundException("High DPI TIFF was not created.", highDpiTiffPath);

        // Get file sizes.
        long defaultSize = new FileInfo(defaultTiffPath).Length;
        long highDpiSize = new FileInfo(highDpiTiffPath).Length;

        // Output the size comparison.
        Console.WriteLine($"Default DPI TIFF size: {defaultSize} bytes");
        Console.WriteLine($"High DPI (500) TIFF size: {highDpiSize} bytes");
    }
}
