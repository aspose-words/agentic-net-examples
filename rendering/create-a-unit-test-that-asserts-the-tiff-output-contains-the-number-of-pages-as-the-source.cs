using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        RunTiffPageCountTest();
        Console.WriteLine("Test completed successfully.");
    }

    // Test that a multi‑page TIFF rendering contains the same number of pages as the source DOCX.
    private static void RunTiffPageCountTest()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Force layout calculation and obtain the page count.
        int sourcePageCount = doc.PageCount;

        // Configure image save options for TIFF output.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff)
        {
            // Optional: set a reasonable resolution.
            Resolution = 300
        };

        // Render each page to a separate TIFF file.
        for (int i = 0; i < sourcePageCount; i++)
        {
            tiffOptions.PageSet = new PageSet(i); // zero‑based page index
            string outPath = Path.Combine(artifactsDir, $"output_page_{i + 1}.tiff");
            doc.Save(outPath, tiffOptions);
        }

        // Verify that the number of generated TIFF files matches the source page count.
        string[] tiffFiles = Directory.GetFiles(artifactsDir, "output_page_*.tiff");
        if (tiffFiles.Length != sourcePageCount)
            throw new Exception($"Expected {sourcePageCount} TIFF files, but found {tiffFiles.Length}.");

        // Verify that each TIFF file exists and is not empty.
        foreach (string file in tiffFiles)
        {
            FileInfo info = new FileInfo(file);
            if (!info.Exists)
                throw new Exception($"TIFF file not found: {file}");
            if (info.Length == 0)
                throw new Exception($"TIFF file is empty: {file}");
        }
    }
}
