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

        // Save with default resolution (no explicit DesiredDpi).
        string defaultTiffPath = Path.Combine(outputDir, "Sample_DefaultDpi.tiff");
        ImageSaveOptions defaultOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            UseHighQualityRendering = true,
            UseAntiAliasing = true
        };
        doc.Save(defaultTiffPath, defaultOptions);

        // Save with archival‑quality DPI (500).
        string highDpiTiffPath = Path.Combine(outputDir, "Sample_500Dpi.tiff");
        ImageSaveOptions highDpiOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            Resolution = 500f,               // Desired DPI
            UseHighQualityRendering = true,
            UseAntiAliasing = true
        };
        doc.Save(highDpiTiffPath, highDpiOptions);

        // Verify that both files were created.
        if (!File.Exists(defaultTiffPath) || !File.Exists(highDpiTiffPath))
            throw new FileNotFoundException("One or more TIFF files were not created.");

        // Compare file sizes.
        long defaultSize = new FileInfo(defaultTiffPath).Length;
        long highDpiSize = new FileInfo(highDpiTiffPath).Length;
        long sizeDiff = highDpiSize - defaultSize;
        double percentIncrease = defaultSize > 0 ? (double)sizeDiff / defaultSize * 100 : 0;

        Console.WriteLine($"Default DPI TIFF size: {defaultSize} bytes");
        Console.WriteLine($"500 DPI TIFF size:    {highDpiSize} bytes");
        Console.WriteLine($"Size increase:       {sizeDiff} bytes ({percentIncrease:F2}%)");
    }
}
