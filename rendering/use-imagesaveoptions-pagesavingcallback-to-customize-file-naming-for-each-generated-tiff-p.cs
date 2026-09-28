using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Content of page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 2.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Content of page 3.");

        // Ensure layout is up to date to obtain the correct page count.
        doc.UpdatePageLayout();

        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // Configure ImageSaveOptions for TIFF with a custom page‑saving callback.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        options.PageSavingCallback = new CustomPageSavingCallback(outputFolder);

        // Save the document; the callback will generate separate TIFF files per page.
        string dummyFileName = Path.Combine(outputFolder, "dummy.tiff");
        doc.Save(dummyFileName, options);

        // Verify that each expected page file exists.
        for (int i = 1; i <= doc.PageCount; i++)
        {
            string expectedPath = Path.Combine(outputFolder, $"Page_{i}.tiff");
            if (!File.Exists(expectedPath))
                throw new InvalidOperationException($"Expected page file not found: {expectedPath}");
        }

        Console.WriteLine("TIFF pages saved with custom names.");
    }

    // Callback that sets a custom file name for each rendered page.
    private class CustomPageSavingCallback : IPageSavingCallback
    {
        private readonly string _folder;

        public CustomPageSavingCallback(string folder)
        {
            _folder = folder;
        }

        public void PageSaving(PageSavingArgs args)
        {
            string fileName = Path.Combine(_folder, $"Page_{args.PageIndex + 1}.tiff");
            args.PageFileName = fileName;
        }
    }
}
