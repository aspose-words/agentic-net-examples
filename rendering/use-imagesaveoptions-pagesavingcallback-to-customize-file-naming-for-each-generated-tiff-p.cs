using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Custom callback that sets the file name for each TIFF page.
    private class PageSavingCallback : IPageSavingCallback
    {
        public void PageSaving(PageSavingArgs args)
        {
            // Name each page file as "Page_<page number>.tiff".
            args.PageFileName = $"Page_{args.PageIndex + 1}.tiff";
        }
    }

    public static void Main()
    {
        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 3; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 3)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure TIFF save options with the custom PageSavingCallback.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff);
        saveOptions.PageSavingCallback = new PageSavingCallback();

        // The file name passed here is ignored because the callback provides per‑page names.
        doc.Save("ignored.tiff", saveOptions);

        // Verify that a TIFF file was created for each page.
        int pageCount = doc.PageCount;
        for (int i = 1; i <= pageCount; i++)
        {
            string fileName = $"Page_{i}.tiff";
            if (!File.Exists(fileName))
                throw new FileNotFoundException($"Expected TIFF page file not found: {fileName}");
        }

        Console.WriteLine($"Document rendered to {pageCount} TIFF page files.");
    }
}
