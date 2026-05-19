using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with more than five pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7) // No break after the last page.
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure image save options for TIFF.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Limit rendering to the first five pages (zero‑based indices 0‑4).
        saveOptions.PageSet = new PageSet(new PageRange(0, 4));

        // Save the TIFF file.
        string outputPath = Path.Combine(artifactsDir, "FirstFivePages.tiff");
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("TIFF file was not created.");

        // Optionally, output the result path (no user interaction required).
        Console.WriteLine($"TIFF saved to: {outputPath}");
    }
}
