using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i} of the sample document.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Render with default resolution (no explicit DPI set).
        string defaultTiffPath = Path.Combine(outputDir, "DefaultDpi.tiff");
        ImageSaveOptions defaultOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(defaultTiffPath, defaultOptions);

        // Render with archival‑quality DPI = 500.
        string highDpiTiffPath = Path.Combine(outputDir, "HighDpi_500.tiff");
        ImageSaveOptions highDpiOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 500,                     // Desired DPI.
            UseAntiAliasing = true,               // Improve visual quality.
            UseHighQualityRendering = true        // Enable high‑quality rendering.
        };
        doc.Save(highDpiTiffPath, highDpiOptions);

        // Verify that both files were created.
        if (!File.Exists(defaultTiffPath) || !File.Exists(highDpiTiffPath))
            throw new FileNotFoundException("One or more TIFF files were not generated.");

        // Compare file sizes.
        long defaultSize = new FileInfo(defaultTiffPath).Length;
        long highDpiSize = new FileInfo(highDpiTiffPath).Length;

        Console.WriteLine($"Default DPI TIFF size: {defaultSize} bytes");
        Console.WriteLine($"High DPI (500) TIFF size: {highDpiSize} bytes");
        Console.WriteLine($"Size increase: {highDpiSize - defaultSize} bytes");
    }
}
