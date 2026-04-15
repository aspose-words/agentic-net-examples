using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with more than five pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure image save options for TIFF and limit to the first five pages.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        // Page indices are zero‑based; pages 0‑4 correspond to the first five pages.
        saveOptions.PageSet = new PageSet(new PageRange(0, 4));

        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstFivePages.tiff");
        doc.Save(outputPath, saveOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, report success.
        Console.WriteLine($"TIFF file saved to: {outputPath}");
    }
}
