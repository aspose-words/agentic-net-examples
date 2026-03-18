using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Folder where the generated TIFF pages will be saved.
        string artifactsDir = @"C:\Temp\Artifacts\";
        Directory.CreateDirectory(artifactsDir);

        // Create a simple document with three pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Page 1");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 2");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Page 3");

        // Configure image save options for TIFF output.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
        // Render all pages into a multi‑frame TIFF.
        options.PageLayout = MultiPageLayout.TiffFrames();
        // Assign the custom callback that defines the file name for each page.
        options.PageSavingCallback = new CustomPageSavingCallback(artifactsDir);

        // Save the document. The callback will create separate files for each page.
        doc.Save(Path.Combine(artifactsDir, "output.tiff"), options);
    }

    // Callback that customizes the file name of each saved page.
    private class CustomPageSavingCallback : IPageSavingCallback
    {
        private readonly string _folder;

        public CustomPageSavingCallback(string folder) => _folder = folder;

        public void PageSaving(PageSavingArgs args)
        {
            // PageIndex is zero‑based; add 1 for a more natural numbering.
            string fileName = Path.Combine(_folder, $"MyDocument_Page_{args.PageIndex + 1}.tiff");
            // Set the file name where Aspose.Words will write this page.
            args.PageFileName = fileName;
            // Ensure the stream is closed after the page is written.
            args.KeepPageStreamOpen = false;
        }
    }
}
