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
        builder.Writeln("First page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Second page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Third page.");

        // Configure ImageSaveOptions for TIFF output and assign a custom callback.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.PageSavingCallback = new CustomPageSavingCallback(outputDir);

        // Save the document. The callback will name each page file.
        string dummyFileName = Path.Combine(outputDir, "output.tiff");
        doc.Save(dummyFileName, options);

        // Verify that each page file was created.
        for (int i = 1; i <= doc.PageCount; i++)
        {
            string pagePath = Path.Combine(outputDir, $"Page_{i}.tiff");
            if (!File.Exists(pagePath))
                throw new Exception($"Expected page file not found: {pagePath}");
        }

        Console.WriteLine("TIFF pages saved successfully.");
    }

    // Callback that sets a custom file name for each saved page.
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
            string fileName = Path.Combine(_outputDir, $"Page_{args.PageIndex + 1}.tiff");
            args.PageFileName = fileName;
        }
    }
}
