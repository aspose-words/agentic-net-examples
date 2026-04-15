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
        builder.Writeln("Page 1 – sample text.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2 – more sample text.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3 – final sample text.");

        // Save with the default resolution (96 dpi).
        string defaultPath = Path.Combine(outputDir, "default.tiff");
        ImageSaveOptions defaultOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(defaultPath, defaultOptions);

        // Save with archival‑quality resolution (500 dpi).
        string highDpiPath = Path.Combine(outputDir, "high500.tiff");
        ImageSaveOptions highDpiOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // DesiredDpi is exposed via the Resolution property.
        highDpiOptions.Resolution = 500;
        doc.Save(highDpiPath, highDpiOptions);

        // Verify that both files were created.
        if (!File.Exists(defaultPath) || !File.Exists(highDpiPath))
            throw new InvalidOperationException("TIFF files were not generated as expected.");

        // Compare file sizes.
        long defaultSize = new FileInfo(defaultPath).Length;
        long highSize = new FileInfo(highDpiPath).Length;

        Console.WriteLine($"Default DPI (96) TIFF size: {defaultSize} bytes");
        Console.WriteLine($"High DPI (500) TIFF size: {highSize} bytes");
    }
}
