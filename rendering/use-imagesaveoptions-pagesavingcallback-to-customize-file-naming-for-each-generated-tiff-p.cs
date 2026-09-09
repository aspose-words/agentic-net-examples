using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure ImageSaveOptions for TIFF output and assign a custom callback.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.PageSavingCallback = new CustomPageSavingCallback(outputDir);

        // Save the document; the callback will name each page file.
        string tiffPath = Path.Combine(outputDir, "Document.tiff");
        doc.Save(tiffPath, saveOptions);

        Console.WriteLine($"TIFF pages have been saved to: {outputDir}");
    }

    // Callback that sets a custom file name for each page when saving.
    private class CustomPageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputFolder;

        public CustomPageSavingCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // PageIndex is zero‑based; add 1 for human‑readable numbering.
            string pageFileName = Path.Combine(_outputFolder, $"Page_{args.PageIndex + 1}.tiff");
            args.PageFileName = pageFileName;
            // Keep the default behavior of closing the stream after each page.
        }
    }
}
