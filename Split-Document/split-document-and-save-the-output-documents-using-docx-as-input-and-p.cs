using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocToPdf
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Configure PDF save options to split the document by pages.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Assign a callback that will be invoked for each page.
            PageSavingCallback = new PdfPageSplitter()
        };

        // Save the document. The callback will create separate PDF files for each page.
        doc.Save("Output.pdf", pdfOptions);
    }

    // Callback implementation that defines how each page is saved.
    private class PdfPageSplitter : IPageSavingCallback
    {
        public void PageSaving(PageSavingArgs args)
        {
            // Build a distinct file name for the current page.
            // PageIndex is zero‑based, so add 1 for a more natural numbering.
            string fileName = $"Output_Page_{args.PageIndex + 1}.pdf";

            // Tell Aspose.Words to save this page to the specified file.
            args.PageFileName = fileName;

            // Alternatively you could provide a stream:
            // args.PageStream = new FileStream(fileName, FileMode.Create);
            // args.KeepPageStreamOpen = false;
        }
    }
}
