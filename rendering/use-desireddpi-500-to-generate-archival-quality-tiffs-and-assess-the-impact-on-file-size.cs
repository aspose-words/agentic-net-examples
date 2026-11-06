using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and a builder to add content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add enough text to generate multiple pages.
        builder.Writeln("Sample Document for TIFF Rendering");
        for (int i = 1; i <= 2000; i++)
        {
            builder.Writeln($"This is line {i} of sample text to increase document size and page count.");
        }

        // Define output file names.
        string defaultTiffPath = "output_default.tiff";
        string highDpiTiffPath = "output_500dpi.tiff";

        // Save with default DPI (no explicit resolution set).
        ImageSaveOptions defaultOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(defaultTiffPath, defaultOptions);

        // Save with higher DPI (use Resolution property instead of DesiredDpi).
        ImageSaveOptions highDpiOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 500 // DPI setting for archival‑quality TIFF.
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
        Console.WriteLine($"500 DPI TIFF size: {highDpiSize} bytes");
        Console.WriteLine($"Size increase factor: {(double)highDpiSize / defaultSize:F2}");
    }
}
