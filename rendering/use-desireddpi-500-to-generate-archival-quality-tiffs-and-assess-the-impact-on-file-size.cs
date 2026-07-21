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

        // Render the document to a TIFF using the default DPI (96).
        string defaultTiffPath = Path.Combine(outputDir, "DefaultDpi.tiff");
        ImageSaveOptions defaultOptions = new ImageSaveOptions(SaveFormat.Tiff);
        doc.Save(defaultTiffPath, defaultOptions);

        // Render the same document to a TIFF using a higher DPI (500) for archival quality.
        string highDpiTiffPath = Path.Combine(outputDir, "DesiredDpi500.tiff");
        ImageSaveOptions highDpiOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // The DesiredDpi property does not exist; use the Resolution property instead.
        highDpiOptions.Resolution = 500; // Sets both horizontal and vertical DPI.
        doc.Save(highDpiTiffPath, highDpiOptions);

        // Assess file sizes.
        long defaultSize = new FileInfo(defaultTiffPath).Length;
        long highDpiSize = new FileInfo(highDpiTiffPath).Length;

        Console.WriteLine($"Default DPI TIFF size: {defaultSize} bytes");
        Console.WriteLine($"DesiredDpi 500 TIFF size: {highDpiSize} bytes");
        Console.WriteLine($"Size increase: {highDpiSize - defaultSize} bytes ({(double)highDpiSize / defaultSize:P2} of original)");
    }
}
