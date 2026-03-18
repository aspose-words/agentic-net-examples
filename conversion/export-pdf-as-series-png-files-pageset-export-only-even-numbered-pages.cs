using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportEvenPagesToPng
{
    static void Main()
    {
        // Path to the source PDF document.
        string inputPdfPath = @"C:\Docs\input.pdf";

        // Folder where the PNG files will be written.
        string outputFolder = @"C:\Docs\EvenPagesPng";
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document.
        Document doc = new Document(inputPdfPath);

        // Configure image save options for PNG format.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // Export only the even‑numbered pages (PageSet.Even uses zero‑based indices).
        pngOptions.PageSet = PageSet.Even;

        // Use a callback to name each exported page file.
        pngOptions.PageSavingCallback = new EvenPageSavingCallback(outputFolder);

        // The file name supplied to Save is ignored because the callback provides a name for each page.
        doc.Save(Path.Combine(outputFolder, "placeholder.png"), pngOptions);
    }

    // Callback that assigns a file name to each saved page.
    private class EvenPageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputFolder;

        public EvenPageSavingCallback(string outputFolder) => _outputFolder = outputFolder;

        public void PageSaving(PageSavingArgs args)
        {
            // args.PageIndex is zero‑based within the filtered set (only even pages are processed).
            // Original even pages have odd indices, so calculate the original page number:
            //   0 -> page 2, 1 -> page 4, etc.
            int originalPageNumber = args.PageIndex * 2 + 2;

            string fileName = Path.Combine(_outputFolder, $"Page_{originalPageNumber}.png");
            args.PageFileName = fileName;
        }
    }
}
