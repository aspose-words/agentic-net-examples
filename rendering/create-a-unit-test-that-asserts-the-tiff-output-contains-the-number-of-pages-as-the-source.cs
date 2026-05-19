using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    // Callback that counts how many pages are saved during the TIFF rendering process.
    private class PageCountingCallback : IPageSavingCallback
    {
        public int SavedPageCount { get; private set; }

        public void PageSaving(PageSavingArgs args)
        {
            // Increment the counter each time a page is saved.
            SavedPageCount++;
        }
    }

    public static void Main()
    {
        // Prepare output directory.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Paths for the source DOCX and the rendered TIFF.
        string docPath = Path.Combine(artifactsDir, "Sample.docx");
        string tiffPath = Path.Combine(artifactsDir, "Sample.tiff");

        // Create a sample document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Save the source document (optional, just for reference).
        doc.Save(docPath);

        // Get the page count of the source document.
        int sourcePageCount = doc.PageCount;

        // Set up TIFF save options to render all pages into a multi‑frame TIFF.
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        tiffOptions.PageLayout = MultiPageLayout.TiffFrames();

        // Attach the callback that will count saved pages.
        PageCountingCallback callback = new PageCountingCallback();
        tiffOptions.PageSavingCallback = callback;

        // Render the document to TIFF.
        doc.Save(tiffPath, tiffOptions);

        // Verify that the TIFF file was created.
        if (!File.Exists(tiffPath))
            throw new Exception("TIFF file was not created.");

        // Verify that the number of pages saved matches the source document's page count.
        if (callback.SavedPageCount != sourcePageCount)
            throw new Exception($"TIFF page count ({callback.SavedPageCount}) does not match source document page count ({sourcePageCount}).");

        Console.WriteLine("TIFF page count matches the source DOCX page count.");
    }
}
