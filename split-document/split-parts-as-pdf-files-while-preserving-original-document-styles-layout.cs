using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace SplitPdfExample
{
    // Callback that saves each page of the document as an individual PDF file.
    class PageSplitCallback : IPageSavingCallback
    {
        private readonly string _outputFolder;

        public PageSplitCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // Build a file name for the current page (page index is zero‑based).
            string pageFileName = Path.Combine(_outputFolder, $"Document_Part_{args.PageIndex + 1}.pdf");

            // Direct Aspose.Words to write this page to the specified file.
            args.PageFileName = pageFileName;

            // No need to keep the stream open after the page is written.
            args.KeepPageStreamOpen = false;
        }
    }

    class Program
    {
        static void Main()
        {
            // Input Word document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Folder where the split PDF parts will be stored.
            string outputFolder = @"C:\Docs\SplitPdfParts";

            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Load the original document (preserves all styles and layout).
            Document doc = new Document(inputPath);

            // Configure PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use a callback to write each page to a separate PDF file.
                PageSavingCallback = new PageSplitCallback(outputFolder)
            };

            // The main file name is required but will not contain any pages because
            // the callback redirects each page to its own file.
            string dummyMainFile = Path.Combine(outputFolder, "Dummy.pdf");

            // Save the document; the callback handles the actual per‑page files.
            doc.Save(dummyMainFile, pdfOptions);
        }
    }
}
