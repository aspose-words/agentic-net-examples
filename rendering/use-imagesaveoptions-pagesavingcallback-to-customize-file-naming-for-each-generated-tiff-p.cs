using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define output folder.
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

        // Configure ImageSaveOptions for TIFF and assign a custom PageSavingCallback.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.PageSavingCallback = new CustomPageSavingCallback(outputDir);

        // Save the document. The callback will create separate TIFF files for each page.
        string dummyFileName = Path.Combine(outputDir, "dummy.tiff"); // Name is overridden by the callback.
        doc.Save(dummyFileName, saveOptions);

        // Verify that each expected page file was created.
        for (int i = 0; i < doc.PageCount; i++)
        {
            string expectedPath = Path.Combine(outputDir, $"Page_{i}.tiff");
            if (!File.Exists(expectedPath))
                throw new FileNotFoundException($"Expected page file not found: {expectedPath}");
        }

        Console.WriteLine("TIFF pages saved successfully.");
    }

    // Callback that customizes the file name for each saved page.
    private class CustomPageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputDir;

        public CustomPageSavingCallback(string outputDir)
        {
            _outputDir = outputDir;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // Generate a custom file name for the current page.
            string pageFileName = Path.Combine(_outputDir, $"Page_{args.PageIndex}.tiff");
            args.PageFileName = pageFileName;

            // Use a stream to write the page data.
            args.PageStream = new FileStream(pageFileName, FileMode.Create);
            args.KeepPageStreamOpen = false;
        }
    }
}
