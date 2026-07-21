using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new document and add several pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add 7 pages of sample text.
        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure image save options for TIFF format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);

        // Limit the output to the first five pages (zero‑based indices 0‑4).
        saveOptions.PageSet = new PageSet(new PageRange(0, 4));

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FirstFivePages.tiff");

        // Save the document as a multi‑page TIFF using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The TIFF file was not created.");

        // Optionally, report success.
        Console.WriteLine($"TIFF file saved successfully to: {outputPath}");
    }
}
