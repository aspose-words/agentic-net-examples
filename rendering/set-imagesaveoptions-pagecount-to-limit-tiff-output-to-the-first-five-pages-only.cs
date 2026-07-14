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

        // Create a sample document with more than five pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7) // No break after the last page.
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure image save options for TIFF and limit to the first five pages.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Page indices are zero‑based; pages 0‑4 correspond to the first five pages.
        saveOptions.PageSet = new PageSet(0, 1, 2, 3, 4);

        // Save the TIFF file.
        string tiffPath = Path.Combine(outputDir, "FirstFivePages.tiff");
        doc.Save(tiffPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(tiffPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Optionally, report success.
        Console.WriteLine($"TIFF saved to: {tiffPath}");
    }
}
