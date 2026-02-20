using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentPageSplitter
{
    // Implements a callback that names each saved page file.
    public class CustomPageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputFolder;

        public CustomPageSavingCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
        }

        public void PageSaving(PageSavingArgs args)
        {
            // Build a file name like "Page_0.html", "Page_1.html", etc.
            string fileName = Path.Combine(_outputFolder, $"Page_{args.PageIndex}.html");

            // Option 1: set the file name directly.
            args.PageFileName = fileName;

            // Option 2: provide a stream (optional, shown for completeness).
            // args.PageStream = new FileStream(fileName, FileMode.Create);

            // Ensure the stream is closed after saving.
            args.KeepPageStreamOpen = false;
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Path to the source DOCX document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Folder where individual page files will be written.
            string outputFolder = @"C:\Docs\SplitPages";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Configure HTML fixed save options to split pages.
            HtmlFixedSaveOptions saveOptions = new HtmlFixedSaveOptions
            {
                // Assign the callback that will name each page file.
                PageSavingCallback = new CustomPageSavingCallback(outputFolder)
            };

            // The base file name is not used when PageSavingCallback sets the file name,
            // but a valid path is still required.
            string dummyOutputPath = Path.Combine(outputFolder, "dummy.html");

            // Save the document; each page will be saved as a separate HTML file.
            doc.Save(dummyOutputPath, saveOptions);
        }
    }
}
