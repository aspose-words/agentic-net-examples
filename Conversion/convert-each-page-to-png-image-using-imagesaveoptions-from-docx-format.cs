using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX document.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Directory where the PNG images will be saved.
        string outputDir = @"C:\Docs\PagesAsPng\";
        Directory.CreateDirectory(outputDir);

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure image save options for PNG format.
        ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render all pages (default). No need to set PageSet.
            // Use a callback to control how each page is saved.
            PageSavingCallback = new PageSavingCallback(outputDir)
        };

        // Save the document; the callback will create one PNG per page.
        doc.Save(Path.Combine(outputDir, "placeholder.png"), saveOptions);
    }

    // Callback that saves each rendered page to a separate PNG file.
    private class PageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputFolder;

        public PageSavingCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // Build a file name like "Page_0.png", "Page_1.png", etc.
            string fileName = $"Page_{args.PageIndex}.png";

            // Set the full path for the page image.
            args.PageFileName = Path.Combine(_outputFolder, fileName);

            // Ensure the stream is closed after saving.
            args.KeepPageStreamOpen = false;
        }
    }
}
