using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a multi‑page document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add content for three pages.
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Define save options for TIFF and limit rendering to the first page (zero‑based index).
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.PageSet = new PageSet(0); // Render only the first page.

        string outputPath = "FirstPage.tiff";

        // Save the first page as a single‑page TIFF.
        doc.Save(outputPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create TIFF file at '{outputPath}'.");

        // Optionally, report success (no interactive prompts required).
        Console.WriteLine($"TIFF file created successfully: {Path.GetFullPath(outputPath)}");
    }
}
