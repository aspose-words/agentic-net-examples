using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace DocumentToJpegPages
{
    // Callback that defines how each page is saved as a separate JPEG file.
    public class CustomPageSavingCallback : IPageSavingCallback
    {
        private readonly string _outputFolder;

        public CustomPageSavingCallback(string outputFolder)
        {
            _outputFolder = outputFolder;
            // Ensure the output directory exists.
            Directory.CreateDirectory(_outputFolder);
        }

        public void PageSaving(PageSavingArgs args)
        {
            // Build a file name like "Page_1.jpg", "Page_2.jpg", etc.
            string fileName = $"Page_{args.PageIndex + 1}.jpg";
            args.PageFileName = Path.Combine(_outputFolder, fileName);

            // Use a file stream for the page output.
            args.PageStream = new FileStream(args.PageFileName, FileMode.Create);
            args.KeepPageStreamOpen = false; // Let Aspose close the stream after saving.
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source DOCX file (10 pages).
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Folder where the JPEG images will be saved.
            string outputFolder = @"C:\Docs\PagesAsJpeg\";

            // Load the document.
            Document doc = new Document(inputPath);

            // Configure image save options for JPEG format.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Jpeg)
            {
                // Use the callback to save each page separately.
                PageSavingCallback = new CustomPageSavingCallback(outputFolder)
            };

            // The file name supplied here is ignored because the callback provides per‑page names.
            doc.Save(Path.Combine(outputFolder, "placeholder.jpg"), saveOptions);
        }
    }
}
