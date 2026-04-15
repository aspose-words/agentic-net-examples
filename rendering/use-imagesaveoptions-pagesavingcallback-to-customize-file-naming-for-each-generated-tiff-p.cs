using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output directory and ensure it exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3.");

        // Configure ImageSaveOptions for TIFF and assign a custom page‑saving callback.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.PageSavingCallback = new CustomPageSavingCallback(outputDir);

        // Save the document. The callback will be invoked for each page.
        string combinedTiffPath = Path.Combine(outputDir, "Combined.tiff");
        doc.Save(combinedTiffPath, saveOptions);

        // Verify that a separate TIFF file was created for each page.
        string[] pageFiles = Directory.GetFiles(outputDir, "Page_*.tiff");
        if (pageFiles.Length != doc.PageCount)
            throw new InvalidOperationException($"Expected {doc.PageCount} page files, but found {pageFiles.Length}.");

        // Indicate successful completion.
        Console.WriteLine("TIFF pages saved:");
        foreach (string file in pageFiles)
            Console.WriteLine(file);
    }

    // Callback that customizes the file name (and optionally the stream) for each saved page.
    private class CustomPageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputDir;

        public CustomPageSavingCallback(string outputDir)
        {
            _outputDir = outputDir;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // PageIndex is zero‑based; add 1 for human‑readable numbering.
            string pageFileName = Path.Combine(_outputDir, $"Page_{args.PageIndex + 1}.tiff");
            args.PageFileName = pageFileName;

            // Optionally, you could provide a stream instead:
            // args.PageStream = new FileStream(pageFileName, FileMode.Create);
        }
    }
}
