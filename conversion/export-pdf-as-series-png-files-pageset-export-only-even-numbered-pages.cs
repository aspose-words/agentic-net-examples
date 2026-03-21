using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class ExportEvenPagesToPng
{
    static void Main()
    {
        // Create a temporary PDF with a few pages.
        string tempPdfPath = Path.Combine(Path.GetTempPath(), "temp_input.pdf");
        CreateSamplePdf(tempPdfPath, 5); // 5 pages

        // Folder where the PNG files will be written.
        string outputFolder = Path.Combine(Path.GetTempPath(), "EvenPagesPng");
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document.
        Document doc = new Document(tempPdfPath);

        // Configure image save options for PNG format.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Export only the even‑numbered pages (PageSet.Even uses zero‑based indices).
            PageSet = PageSet.Even,
            PageSavingCallback = new EvenPageSavingCallback(outputFolder)
        };

        // The file name supplied to Save is ignored because the callback provides a name for each page.
        doc.Save(Path.Combine(outputFolder, "placeholder.png"), pngOptions);

        Console.WriteLine($"Exported even pages to PNG files in: {outputFolder}");
    }

    // Generates a simple multi‑page PDF for demonstration purposes.
    private static void CreateSamplePdf(string filePath, int pageCount)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < pageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(filePath);
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
