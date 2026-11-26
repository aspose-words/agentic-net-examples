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
        builder.Writeln("Content of the first page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of the second page.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of the third page.");

        // Configure TIFF save options with a custom page-saving callback.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.PageSavingCallback = new CustomPageSavingCallback(outputDir);
        // Optional: set additional options such as compression.
        saveOptions.TiffCompression = TiffCompression.Ccitt4;

        // Save the document. The callback will generate separate TIFF files per page.
        string dummyFileName = Path.Combine(outputDir, "Combined.tiff");
        doc.Save(dummyFileName, saveOptions);

        // Verify that each expected TIFF page file was created.
        int pageCount = doc.PageCount;
        for (int i = 0; i < pageCount; i++)
        {
            string expectedPath = Path.Combine(outputDir, $"Page_{i + 1}.tiff");
            if (!File.Exists(expectedPath))
                throw new Exception($"Expected TIFF page file not found: {expectedPath}");
        }

        // Indicate successful completion (optional).
        Console.WriteLine($"Successfully generated {pageCount} TIFF page files in '{outputDir}'.");
    }

    // Implementation of the page-saving callback to customize file names.
    private class CustomPageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputDirectory;

        public CustomPageSavingCallback(string outputDirectory)
        {
            _outputDirectory = outputDirectory;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // PageIndex is zero‑based; create a file name like "Page_1.tiff".
            string fileName = $"Page_{args.PageIndex + 1}.tiff";
            args.PageFileName = Path.Combine(_outputDirectory, fileName);
        }
    }
}
