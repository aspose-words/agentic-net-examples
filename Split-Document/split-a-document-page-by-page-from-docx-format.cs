using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the source DOCX document.
        Document doc = new Document("Input.docx");

        // Configure HTML fixed save options to split the document page by page.
        HtmlFixedSaveOptions htmlOptions = new HtmlFixedSaveOptions();

        // Assign a callback that will name each page file.
        htmlOptions.PageSavingCallback = new PageSavingCallback();

        // Save the document. The callback will create separate HTML files for each page.
        doc.Save("CombinedOutput.html", htmlOptions);
    }

    // Callback implementation that defines how each page is saved.
    private class PageSavingCallback : IPageSavingCallback
    {
        public void PageSaving(PageSavingArgs args)
        {
            // Build a file name that includes the zero‑based page index.
            string pageFileName = $"Page_{args.PageIndex}.html";

            // Option 1: set the file name directly.
            args.PageFileName = pageFileName;

            // Option 2 (alternative): write to a custom stream.
            // args.PageStream = new FileStream(pageFileName, FileMode.Create);

            // Ensure the stream is closed after saving.
            args.KeepPageStreamOpen = false;
        }
    }
}
